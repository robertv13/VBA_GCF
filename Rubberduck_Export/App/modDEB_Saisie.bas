Attribute VB_Name = "modDEB_Saisie"
'@Folder("Saisie_Déboursé")

Option Explicit

'Variables globales
Public sauvegardesCaracteristiquesForme As Object
Public numeroDebourseARenverser As Long

Sub shp_DEB_Saisie_Update_Click()

    Call DEB_Saisie_Update

End Sub

Sub DEB_Saisie_Update()

    If wshDEB_Saisie.Range("B7").Value = True Then
        Call DEB_Renversement_Update
        Exit Sub
    End If
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Saisie_Update", "", 0)
    
    'Remove highlight from last cell
    If wshDEB_Saisie.Range("B4").Value <> "" Then
        wshDEB_Saisie.Range(wshDEB_Saisie.Range("B4").Value).Interior.Color = xlNone
    End If
    
    'Date is not valid OR the transaction does not balance
    If Fn_Is_Date_Valide(wshDEB_Saisie.Range("O4").Value) = False Or _
        Fn_Is_Debours_Balance = False Then
            Exit Sub
    End If
    
    'Is every line of the transaction well entered ?
    Dim rowDebSaisie As Long
    rowDebSaisie = wshDEB_Saisie.Range("E23").End(xlUp).Row  'Last Used Row in wshDEB_Saisie
    If Fn_Is_Deb_Saisie_Valid(rowDebSaisie) = False Then Exit Sub
    
    'Get the FournID
    wshDEB_Saisie.Range("B5").Value = Fn_GetID_From_Fourn_Name(wshDEB_Saisie.Range("J4").Value)

    'Transfert des données vers DEB_Trans
    Call DEB_Trans_Add_Record_To_DB(rowDebSaisie)
    Call DEB_Trans_Add_Record_Locally(rowDebSaisie)
    
    'GL posting
    Call DEB_Saisie_GL_Posting_Preparation
    
    If wshDEB_Saisie.ckbRecurrente = True Then
        Call Save_DEB_Recurrent(rowDebSaisie)
    End If
    
    'Retrieve the CurrentDebours number
    Dim CurrentDeboursNo As String
    CurrentDeboursNo = wshDEB_Saisie.Range("B1").Value
    
    MsgBox "Le déboursé, numéro '" & CurrentDeboursNo & "' a été reporté avec succès"
    
    'Get ready for a new one
    Call DEB_Saisie_Clear_All_Cells
    
    Application.EnableEvents = True
    
    wshDEB_Saisie.Activate
    wshDEB_Saisie.Range("F4").Select
        
    Call Log_Record("modDEB_Saisie:DEB_Saisie_Update", "", startTime)
        
End Sub

Sub DEB_Renversement_Update()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Renversement_Update", "", 0)
    
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
    ws.Range("B5").Value = Fn_GetID_From_Fourn_Name(wshDEB_Saisie.Range("J4").Value)
    
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
    Call DEB_Trans_Add_Record_To_DB(rowLastUsed)
    Call DEB_Trans_Add_Record_Locally(rowLastUsed)
    
    'Mettre à jour le débouré renversé
    Call DEB_Trans_MAJ_Debourse_Renverse_To_DB
    Call DEB_Trans_MAJ_Debourse_Renverse_Locally
    
    'GL posting
    Call DEB_Saisie_GL_Posting_Preparation
    
    MsgBox "Le déboursé a été RENVERSÉ avec succès", vbInformation, "Confirmation de traitement"
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    DoEvents
    
    'Reorganise wsdDEB_Trans
    Application.ScreenUpdating = False
    Dim shp As Shape
    Set shp = ws.Shapes("btnUpdate")
    Call DEB_Forme_Restaurer(shp)
    
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
    numeroDebourseARenverser = -1
    ws.Range("B7").Value = False

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set shp = Nothing
    Set ws = Nothing
    
    Call Log_Record("modDEB_Saisie:DEB_Renversement_Update", "", startTime)
    
End Sub

Sub DEB_Trans_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_To_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "DEB_Trans$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(NoEntrée) AS MaxDebTransNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastDebTrans As Long
    If IsNull(rs.Fields("MaxDebTransNo").Value) Then
        'Handle empty table (assign a default value, e.g., 0)
        lastDebTrans = 0
    Else
        lastDebTrans = rs.Fields("MaxDebTransNo").Value
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
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Read all line from Journal Entry
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            With wshDEB_Saisie
                'Add fields to the recordset before updating it
                rs.Fields(fDebTNoEntrée - 1).Value = currDebTransNo
                rs.Fields(fDebTDate - 1).Value = .Range("O4").Value
                rs.Fields(fDebTType - 1).Value = .Range("F4").Value
                rs.Fields(fDebTBeneficiaire - 1).Value = .Range("J4").Value
                rs.Fields(fDebTFournID - 1).Value = .Range("B5").Value
                rs.Fields(fDebTDescription - 1).Value = .Range("F6").Value & IIf(.Range("B7"), " (RENVERSEMENT de " & numeroDebourseARenverser & ")", "")
                rs.Fields(fDebTReference - 1).Value = .Range("M6").Value
                
                rs.Fields(fDebTNoCompte - 1).Value = .Range("Q" & l).Value
                rs.Fields(fDebTCompte - 1).Value = .Range("E" & l).Value
                rs.Fields(fDebTCodeTaxe - 1).Value = .Range("H" & l).Value
                rs.Fields(fDebTTotal - 1).Value = CDbl(.Range("I" & l).Value)
                rs.Fields(fDebTTPS - 1).Value = CDbl(.Range("J" & l).Value)
                rs.Fields(fDebTTVQ - 1).Value = CDbl(.Range("K" & l).Value)
                rs.Fields(fDebTCréditTPS - 1).Value = CDbl(.Range("L" & l).Value)
                rs.Fields(fDebTCréditTVQ - 1).Value = CDbl(.Range("M" & l).Value)
                'Montant de dépense (Total - creditTPS - creditTVQ)
                rs.Fields(fDebTDépense - 1).Value = CDbl(.Range("I" & l).Value _
                                                  - .Range("L" & l).Value _
                                                  - .Range("M" & l).Value)
                rs.Fields(fDebTAutreRemarque - 1).Value = ""
                rs.Fields(fDebTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            End With
        rs.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_To_DB", "", startTime)

End Sub

Sub DEB_Trans_Add_Record_Locally(r As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("*** modDEB_Saisie:DEB_Trans_Add_Record_Locally", CStr(r), 0)
    
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
            ws.Cells(rowToBeUsed, fDebTDescription).Value = .Range("F6").Value & IIf(.Range("B7"), " (RENVERSEMENT de " & numeroDebourseARenverser & ")", "")
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
            ws.Cells(rowToBeUsed, fDebTAutreRemarque).Value = ""
            ws.Cells(rowToBeUsed, fDebTTimeStamp).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        End With
        rowToBeUsed = rowToBeUsed + 1
        Call Log_Record("    modDEB_Saisie:DEB_Trans_Add_Record_Locally", -1)
    Next i
    
    Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_Locally", "", startTime)

    Application.ScreenUpdating = True

End Sub

Sub DEB_Trans_MAJ_Debourse_Renverse_To_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Trans_MAJ_Debourse_Renverse_To_DB", "", 0)
    
    'Définition des paramètres
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "DEB_Trans$"

    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Requête SQL pour rechercher la ligne correspondante
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE [NoEntrée] = " & numeroDebourseARenverser

    'Ouvrir le Recordset
    rs.Open strSQL, conn, 1, 3 'adOpenKeyset (1) + adLockOptimistic (3) pour modifier les données

    'Vérifier si des enregistrements existent
    If rs.EOF Then
        MsgBox "Aucun enregistrement trouvé.", vbCritical, "Impossible de mettre à jour les déboursés RENVERSÉS"
    Else
        'Boucler à travers les enregistrements
        Do While Not rs.EOF
        ' Vérifier si Reference contient déjà "RENVERSÉ" pour éviter les doublons
        If InStr(1, rs.Fields(fDebTDescription - 1).Value, " (RENVERSÉ", vbTextCompare) = 0 Then
            rs.Fields(fDebTDescription - 1).Value = rs.Fields(fDebTDescription - 1).Value & " (RENVERSÉ par " & wshDEB_Saisie.Range("B1").Value & ")"
            rs.Update
        End If
        'Passer à l'enregistrement suivant
        rs.MoveNext
        Loop
    End If
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing

    Call Log_Record("modDEB_Saisie:DEB_Trans_MAJ_Debourse_Renverse_To_DB", "", startTime)
    
End Sub

Sub DEB_Trans_MAJ_Debourse_Renverse_Locally()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Trans_MAJ_Debourse_Renverse_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = wsdDEB_Trans
    
    'Dernière ligne de la table
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    'Boucler sur toutes les lignes pour trouver les correspondances
    Dim cell As Range
    For Each cell In ws.Range("A2:A" & lastUsedRow)
        If cell.Value = numeroDebourseARenverser Then
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

    Call Log_Record("modDEB_Saisie:DEB_Trans_MAJ_Debourse_Renverse_Locally", "", startTime)
    
End Sub

Sub DEB_AfficherDeboursRecurrent()

    ufListeDEBAuto.show

End Sub

Sub Preparer_Liste_Debourses_Pour_Afficher()

    'Afficher le UserForm
    ufListeDebourse.show vbModal
    
    If numeroDebourseARenverser = -1 Then
        wshDEB_Saisie.Range("F4").Value = ""
        wshDEB_Saisie.Range("F4").Select
    Else
        wshDEB_Saisie.Range("B7").Value = True
    End If
    
End Sub

Sub DEB_Renverser_Ecriture() '2025-02-23 @ 16:56

    Dim ws As Worksheet: Set ws = wsdDEB_Trans
    
    '1. Quelle écriture doit-on renverser (à partir d'un ListBox)
    Call Preparer_Liste_Debourses_Pour_Afficher
    
    If numeroDebourseARenverser = -1 Then
        MsgBox "Vous n'avez sélectionné aucun déboursé à renverser", vbInformation, "Sélection d'un déboursé à renverser"
        Application.EnableEvents = True
        wshDEB_Saisie.Range("F4").Value = ""
        wshDEB_Saisie.Range("F4").Select
        Application.EnableEvents = False
        GoTo Nettoyage
    End If
    
    '2. Aller chercher les debourses pour le numero choisi (0 à n lignes)
    Dim debTransSubset As Variant
    debTransSubset = RechercherLignesTableau(ws, numeroDebourseARenverser)
    
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
    Set shp = wshDEB_Saisie.Shapes("btnUpdate")
    Call DEB_Forme_Modifier(shp)

Nettoyage:

    'Libérer la mémoire
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub DEB_Saisie_GL_Posting_Preparation() '2024-06-05 @ 18:28

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Saisie_GL_Posting_Preparation", "", 0)

    Dim Montant As Double, dateDebours As Date
    Dim descGL_Trans As String, Source As String, deboursType As String
    Dim GL_TransNo As Long
    
    dateDebours = wshDEB_Saisie.Range("O4").Value
    deboursType = wshDEB_Saisie.Range("F4").Value
    descGL_Trans = deboursType & " - " & wshDEB_Saisie.Range("F6").Value
    If Trim$(wshDEB_Saisie.Range("M6").Value) <> "" Then
        descGL_Trans = descGL_Trans & " [" & wshDEB_Saisie.Range("M6").Value & "]"
    End If
    If wshDEB_Saisie.Range("B7").Value = False Then
        Source = "DÉBOURSÉ:" & Format$(wshDEB_Saisie.Range("B1").Value, "00000")
    Else
        Source = "RENV/DÉBOURSÉ:" & Format$(numeroDebourseARenverser, "00000")
    End If
    
    Dim MyArray() As String
    ReDim MyArray(1 To 16, 1 To 4)
    
    'Based on Disbursement type, the CREDIT account will be different
    'Disbursement Total (wshDEB_Saisie.Range("O6"))
    Montant = wshDEB_Saisie.Range("O6").Value
    
    Dim GLNo_Credit As String
    
    Select Case deboursType
        Case "Chèque", "Virement", "Paiement pré-autorisé"
            MyArray(1, 1) = ObtenirNoGlIndicateur("Encaisse")
            MyArray(1, 2) = "Encaisse"
        Case "Carte de crédit"
            MyArray(1, 1) = ObtenirNoGlIndicateur("Carte de crédit")
            MyArray(1, 2) = "Carte de crédit"
        Case "Avances avec Guillaume Charron"
            MyArray(1, 1) = ObtenirNoGlIndicateur("Avances Guillaume Charron")
            MyArray(1, 2) = "Avances avec Guillaume Charron"
        Case "Avances avec 9249-3626 Québec inc."
            MyArray(1, 1) = ObtenirNoGlIndicateur("Avances 9249-3626 Québec inc.")
            MyArray(1, 2) = "Avances avec 9249-3626 Québec inc."
        Case "Avances avec 9333-4829 Québec inc."
            MyArray(1, 1) = ObtenirNoGlIndicateur("Avances 9333-4829 Québec inc.")
            MyArray(1, 2) = "Avances avec 9333-4829 Québec inc."
        Case "Autre"
            MyArray(1, 1) = ObtenirNoGlIndicateur("Encaisse")
            MyArray(1, 2) = "Encaisse"
        Case Else
            MyArray(1, 1) = ObtenirNoGlIndicateur("Encaisse")
            MyArray(1, 2) = "Encaisse"
    End Select
    
    MyArray(1, 3) = -Montant
    MyArray(1, 4) = ""
    
    'Process every lines
    Dim lastUsedRow As Long
    lastUsedRow = wshDEB_Saisie.Cells(wshDEB_Saisie.Rows.count, "E").End(xlUp).Row

    Dim l As Long, arrRow As Long
    arrRow = 2 '1 is already used
    For l = 9 To lastUsedRow
        MyArray(arrRow, 1) = wshDEB_Saisie.Range("Q" & l).Value
        MyArray(arrRow, 2) = wshDEB_Saisie.Range("E" & l).Value
        MyArray(arrRow, 3) = wshDEB_Saisie.Range("N" & l).Value
        MyArray(arrRow, 4) = ""
        arrRow = arrRow + 1
        
        If wshDEB_Saisie.Range("L" & l).Value <> 0 Then
            MyArray(arrRow, 1) = ObtenirNoGlIndicateur("TPS Payée")
            MyArray(arrRow, 2) = "TPS payées"
            MyArray(arrRow, 3) = wshDEB_Saisie.Range("L" & l).Value
            MyArray(arrRow, 4) = ""
            arrRow = arrRow + 1
        End If

        If wshDEB_Saisie.Range("M" & l).Value <> 0 Then
            MyArray(arrRow, 1) = ObtenirNoGlIndicateur("TVQ Payée")
            MyArray(arrRow, 2) = "TVQ payées"
            MyArray(arrRow, 3) = wshDEB_Saisie.Range("M" & l).Value
            MyArray(arrRow, 4) = ""
            arrRow = arrRow + 1
        End If
    Next l
    
    Dim GLEntryNo As Long
    Call GL_Posting_To_DB(dateDebours, descGL_Trans, Source, MyArray, GLEntryNo)
    
    Call GL_Posting_Locally(dateDebours, descGL_Trans, Source, MyArray, GLEntryNo)
    
    Call Log_Record("modDEB_Saisie:DEB_Saisie_GL_Posting_Preparation", "", startTime)

End Sub

Sub ChargerDEBRecurrentDansSaisie(DEBAutoDesc As String, noDEBAuto As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:ChargerDEBRecurrentDansSaisie", "", 0)
    
    'On copie l'écriture automatique vers wshDEB_Saisie
    Dim rowDEBAuto As Long, rowDEB As Long
    rowDEBAuto = wsdDEB_Recurrent.Cells(wsdDEB_Recurrent.Rows.count, "C").End(xlUp).Row  'Last Row used in wshDEB_Recurrent
    
    Call DEB_Saisie_Clear_All_Cells
    
    rowDEB = 9
    
    Application.EnableEvents = False
    Dim r As Long, totAmount As Currency, typeDEB As String
    For r = 2 To rowDEBAuto
        If wsdDEB_Recurrent.Range("A" & r).Value = noDEBAuto And wsdDEB_Recurrent.Range("F" & r).Value <> "" Then
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
            If typeDEB = "" Then
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

    Call Log_Record("modDEB_Saisie:ChargerDEBRecurrentDansSaisie", "", startTime)
    
End Sub

Sub Save_DEB_Recurrent(ll As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:Save_DEB_Recurrent", "", 0)
    
    Dim rowDEBLast As Long
    rowDEBLast = wshDEB_Saisie.Cells(wshDEB_Saisie.Rows.count, "E").End(xlUp).Row  'Last Used Row in wshDEB_Saisie
    
    Call DEB_Recurrent_Add_Record_To_DB(rowDEBLast)
    Call DEB_Recurrent_Add_Record_Locally(rowDEBLast)
    
    Call Log_Record("modDEB_Saisie:Save_DEB_Recurrent", "", startTime)
    
End Sub

Sub DEB_Recurrent_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Recurrent_Add_Record_To_DB", "", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "DEB_Récurrent$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxDebRecNo As Long
    strSQL = "SELECT MAX(NoDebRec) AS MaxDebRecNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastDR As Long, nextDRNo As Long
    If IsNull(rs.Fields("MaxDebRecNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastDR = 0
    Else
        lastDR = rs.Fields("MaxDebRecNo").Value
    End If
    
    'Calculate the new ID
    nextDRNo = lastDR + 1
    wshDEB_Saisie.Range("B2").Value = nextDRNo

    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            With wshDEB_Saisie
                'Add fields to the recordset before updating it
                rs.Fields(fDebRNoDebRec - 1).Value = nextDRNo
                rs.Fields(fDebRDate - 1).Value = .Range("O4").Value
                rs.Fields(fDebRType - 1).Value = .Range("F4").Value
                rs.Fields(fDebRBeneficiaire - 1).Value = .Range("J4").Value
                rs.Fields(fDebRReference - 1).Value = .Range("M6").Value
                rs.Fields(fDebRNoCompte - 1).Value = .Range("Q" & l).Value
                rs.Fields(fDebRCompte - 1).Value = .Range("E" & l).Value
                rs.Fields(fDebRCodeTaxe - 1).Value = .Range("H" & l).Value
                rs.Fields(fDebRTotal - 1).Value = .Range("I" & l).Value
                rs.Fields(fDebRTPS - 1).Value = .Range("J" & l).Value
                rs.Fields(fDebRTVQ - 1).Value = .Range("K" & l).Value
                rs.Fields(fDebRCréditTPS - 1).Value = .Range("L" & l).Value
                rs.Fields(fDebRCréditTVQ - 1).Value = .Range("M" & l).Value
                rs.Fields(fDebRTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            End With
        rs.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modDEB_Saisie:DEB_Recurrent_Add_Record_To_DB", "", startTime)

End Sub

Sub DEB_Recurrent_Add_Record_Locally(r As Long) 'Write records to local file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Recurrent_Add_Record_Locally", "", 0)
    
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
    
    Call DEB_Recurrent_Build_Summary '2024-03-14 @ 07:40
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modDEB_Saisie:DEB_Recurrent_Add_Record_Locally", "", startTime)
    
End Sub

Sub DEB_Forme_Modifier(forme As Shape)

    'Appliquer des modifications à la forme
    Application.ScreenUpdating = True
    forme.Fill.ForeColor.RGB = RGB(255, 0, 0)  ' Rouge
    forme.Line.ForeColor.RGB = RGB(255, 255, 255) ' Noir
    forme.TextFrame2.TextRange.Text = "Renversement"
    forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    DoEvents
    
End Sub

Sub DEB_Recurrent_Build_Summary()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Recurrent_Build_Summary", "", 0)
    
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

    Call Log_Record("modDEB_Saisie:DEB_Recurrent_Build_Summary", "", startTime)

End Sub

Public Sub DEB_Saisie_Clear_All_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Saisie_Clear_All_Cells", "", 0)

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
        .pattern = xlNone
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
    
    Call Log_Record("modDEB_Saisie:DEB_Saisie_Clear_All_Cells", "", startTime)

End Sub

Sub shp_DEB_Back_To_Menu_Click()

    Call DEB_Back_To_Menu

End Sub

Sub DEB_Back_To_Menu()
    
    'Rétablir la forme du bouton (Mettre à jour / Renverser)
    Dim shp As Shape
    Set shp = wshDEB_Saisie.Shapes("btnUpdate")
    Call DEB_Forme_Restaurer(shp)

    wshDEB_Saisie.Visible = xlSheetHidden
    
    Application.ScreenUpdating = False
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub Calculate_GST_PST_And_Credits(d As Date, _
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

Sub DEB_Forme_Sauvegarder(forme As Shape)

    ' Vérifier si le Dictionary est déjà instancié, sinon le créer
    If sauvegardesCaracteristiquesForme Is Nothing Then
        Set sauvegardesCaracteristiquesForme = CreateObject("Scripting.Dictionary")
    End If

    ' Sauvegarder les caractéristiques originales de la forme
    sauvegardesCaracteristiquesForme("Left") = forme.Left
    sauvegardesCaracteristiquesForme("Width") = forme.Width
    sauvegardesCaracteristiquesForme("Height") = forme.Height
    sauvegardesCaracteristiquesForme("FillColor") = forme.Fill.ForeColor.RGB
    sauvegardesCaracteristiquesForme("LineColor") = forme.Line.ForeColor.RGB
    sauvegardesCaracteristiquesForme("Text") = forme.TextFrame2.TextRange.Text
    sauvegardesCaracteristiquesForme("TextColor") = forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    
End Sub

Sub DEB_Forme_Restaurer(forme As Shape)

    'Vérifiez si les caractéristiques originales sont sauvegardées
    If sauvegardesCaracteristiquesForme Is Nothing Then
        Exit Sub
    End If

    'Restaurer les caractéristiques de la forme
    forme.Left = sauvegardesCaracteristiquesForme("Left")
    forme.Width = sauvegardesCaracteristiquesForme("Width")
    forme.Height = sauvegardesCaracteristiquesForme("Height")
    forme.Fill.ForeColor.RGB = sauvegardesCaracteristiquesForme("FillColor")
    forme.Line.ForeColor.RGB = sauvegardesCaracteristiquesForme("LineColor")
    forme.TextFrame2.TextRange.Text = sauvegardesCaracteristiquesForme("Text")
    forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = sauvegardesCaracteristiquesForme("TextColor")

End Sub

