Attribute VB_Name = "modDEB_Saisie"
Option Explicit

Sub shp_DEB_Saisie_Update_Click()

    Call DEB_Saisie_Update

End Sub

Sub DEB_Saisie_Update()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Saisie_Update", 0)
    
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
    rowDebSaisie = wshDEB_Saisie.Range("E23").End(xlUp).row  'Last Used Row in wshDEB_Saisie
    If Fn_Is_Deb_Saisie_Valid(rowDebSaisie) = False Then Exit Sub
    
    'Get the FournID
    wshDEB_Saisie.Range("B5").Value = Fn_GetID_From_Fourn_Name(wshDEB_Saisie.Range("J4").Value)

    'Transfert des données vers DEB_Trans, entête d'abord puis une ligne à la fois
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
        
    Call Log_Record("modDEB_Saisie:DEB_Saisie_Update", startTime)
        
End Sub

Sub DEB_Trans_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_To_DB", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
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
    
    'Calculate the new JE number
    Dim currDebTransNo As Long
    currDebTransNo = lastDebTrans + 1
    Application.EnableEvents = False
    wshDEB_Saisie.Range("B1").Value = currDebTransNo
    Application.EnableEvents = True
    
'    'Build formula
'    Dim formula As String
'    formula = "=ROW()"
'
    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Read all line from Journal Entry
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields(fDebTNoEntrée - 1).Value = currDebTransNo
            rs.Fields(fDebTDate - 1).Value = wshDEB_Saisie.Range("O4").Value
            rs.Fields(fDebTType - 1).Value = wshDEB_Saisie.Range("F4").Value
            rs.Fields(fDebTBeneficiaire - 1).Value = wshDEB_Saisie.Range("J4").Value
            rs.Fields(fDebTFournID - 1).Value = wshDEB_Saisie.Range("B5").Value
            rs.Fields(fDebTDescription - 1).Value = wshDEB_Saisie.Range("F6").Value
            rs.Fields(fDebTReference - 1).Value = wshDEB_Saisie.Range("M6").Value
            rs.Fields(fDebTNoCompte - 1).Value = wshDEB_Saisie.Range("Q" & l).Value
            rs.Fields(fDebTCompte - 1).Value = wshDEB_Saisie.Range("E" & l).Value
            rs.Fields(fDebTCodeTaxe - 1).Value = wshDEB_Saisie.Range("H" & l).Value
            rs.Fields(fDebTTotal - 1).Value = CDbl(wshDEB_Saisie.Range("I" & l).Value)
            rs.Fields(fDebTTPS - 1).Value = CDbl(wshDEB_Saisie.Range("J" & l).Value)
            rs.Fields(fDebTTVQ - 1).Value = CDbl(wshDEB_Saisie.Range("K" & l).Value)
            rs.Fields(fDebTCréditTPS - 1).Value = CDbl(wshDEB_Saisie.Range("L" & l).Value)
            rs.Fields(fDebTCréditTVQ - 1).Value = CDbl(wshDEB_Saisie.Range("M" & l).Value)
            'Montant de dépense (Total - creditTPS - creditTVQ)
            rs.Fields(fDebTDépense - 1).Value = CDbl(wshDEB_Saisie.Range("I" & l).Value _
                                              - wshDEB_Saisie.Range("L" & l).Value _
                                              - wshDEB_Saisie.Range("M" & l).Value)
            rs.Fields(fDebTAutreRemarque - 1).Value = ""
            rs.Fields(fDebTTimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
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
    
    Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_To_DB", startTime)

End Sub

Sub DEB_Trans_Add_Record_Locally(r As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("*** modDEB_Saisie:DEB_Trans_Add_Record_Locally(" & r & ")", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentDebTransNo As Long
    currentDebTransNo = wshDEB_Saisie.Range("B1").Value
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshDEB_Trans.Cells(wshDEB_Trans.Rows.count, "A").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Long
    For i = 9 To r
        wshDEB_Trans.Range("A" & rowToBeUsed).Value = currentDebTransNo
        wshDEB_Trans.Range("B" & rowToBeUsed).Value = wshDEB_Saisie.Range("O4").Value
        wshDEB_Trans.Range("C" & rowToBeUsed).Value = wshDEB_Saisie.Range("F4").Value
        wshDEB_Trans.Range("D" & rowToBeUsed).Value = wshDEB_Saisie.Range("J4").Value
        wshDEB_Trans.Range("E" & rowToBeUsed).Value = wshDEB_Saisie.Range("B5").Value
        wshDEB_Trans.Range("F" & rowToBeUsed).Value = wshDEB_Saisie.Range("F6").Value
        wshDEB_Trans.Range("G" & rowToBeUsed).Value = wshDEB_Saisie.Range("M6").Value
        wshDEB_Trans.Range("H" & rowToBeUsed).Value = wshDEB_Saisie.Range("Q" & i).Value
        wshDEB_Trans.Range("I" & rowToBeUsed).Value = wshDEB_Saisie.Range("E" & i).Value
        wshDEB_Trans.Range("J" & rowToBeUsed).Value = wshDEB_Saisie.Range("H" & i).Value
        wshDEB_Trans.Range("K" & rowToBeUsed).Value = wshDEB_Saisie.Range("I" & i).Value
        wshDEB_Trans.Range("L" & rowToBeUsed).Value = wshDEB_Saisie.Range("J" & i).Value
        wshDEB_Trans.Range("M" & rowToBeUsed).Value = wshDEB_Saisie.Range("K" & i).Value
        wshDEB_Trans.Range("N" & rowToBeUsed).Value = wshDEB_Saisie.Range("L" & i).Value
        wshDEB_Trans.Range("O" & rowToBeUsed).Value = wshDEB_Saisie.Range("M" & i).Value
        '$ dépense = Total - creditTPS - creditTVQ
        wshDEB_Trans.Range("P" & rowToBeUsed).Value = wshDEB_Saisie.Range("I" & i).Value _
                                                      - wshDEB_Saisie.Range("L" & i).Value _
                                                      - wshDEB_Saisie.Range("M" & i).Value
        wshDEB_Trans.Range("Q" & rowToBeUsed).Value = ""
        wshDEB_Trans.Range("R" & rowToBeUsed).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
        Call Log_Record("    modDEB_Saisie:DEB_Trans_Add_Record_Locally", -1)
    Next i
    
    Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_Locally", startTime)

    Application.ScreenUpdating = True

End Sub

Sub DEB_Saisie_GL_Posting_Preparation() '2024-06-05 @ 18:28

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Saisie_GL_Posting_Preparation", 0)

    Dim montant As Double, dateDebours As Date
    Dim descGL_Trans As String, source As String, deboursType As String
    Dim GL_TransNo As Long
    
    dateDebours = wshDEB_Saisie.Range("O4").Value
    deboursType = wshDEB_Saisie.Range("F4").Value
    descGL_Trans = deboursType & " - " & wshDEB_Saisie.Range("F6").Value
    If Trim(wshDEB_Saisie.Range("M6").Value) <> "" Then
        descGL_Trans = descGL_Trans & " [" & wshDEB_Saisie.Range("M6").Value & "]"
    End If
    source = "DÉBOURSÉ:" & Format$(wshDEB_Saisie.Range("B1").Value, "00000")
    
    Dim MyArray() As String
    ReDim MyArray(1 To 16, 1 To 4)
    
    'Based on Disbursement type, the CREDIT account will be different
    'Disbursement Total (wshDEB_Saisie.Range("O6"))
    montant = wshDEB_Saisie.Range("O6").Value
    
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
    
    MyArray(1, 3) = -montant
    MyArray(1, 4) = ""
    
    'Process every lines
    Dim lastUsedRow As Long
    lastUsedRow = wshDEB_Saisie.Cells(wshDEB_Saisie.Rows.count, "E").End(xlUp).row

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
    Call GL_Posting_To_DB(dateDebours, descGL_Trans, source, MyArray, GLEntryNo)
    
    Call GL_Posting_Locally(dateDebours, descGL_Trans, source, MyArray, GLEntryNo)
    
    Call Log_Record("modDEB_Saisie:DEB_Saisie_GL_Posting_Preparation", startTime)

End Sub

Sub Load_DEB_Auto_Into_JE(DEBAutoDesc As String, NoDEBAuto As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:Load_DEB_Auto_Into_JE", 0)
    
    'On copie l'écriture automatique vers wshDEB_Saisie
    Dim rowDEBAuto, rowDEB As Long
    rowDEBAuto = wshDEB_Récurrent.Cells(wshDEB_Récurrent.Rows.count, "C").End(xlUp).row  'Last Row used in wshDEB_Recuurent
    
    Call DEB_Saisie_Clear_All_Cells
    
    rowDEB = 9
    
    Application.EnableEvents = False
    Dim r As Long, totAmount As Currency, typeDEB As String
    For r = 2 To rowDEBAuto
        If wshDEB_Récurrent.Range("A" & r).Value = NoDEBAuto And wshDEB_Récurrent.Range("F" & r).Value <> "" Then
            wshDEB_Saisie.Range("E" & rowDEB).Value = wshDEB_Récurrent.Range("G" & r).Value
            wshDEB_Saisie.Range("H" & rowDEB).Value = wshDEB_Récurrent.Range("H" & r).Value
            wshDEB_Saisie.Range("I" & rowDEB).Value = wshDEB_Récurrent.Range("I" & r).Value
            wshDEB_Saisie.Range("J" & rowDEB).Value = wshDEB_Récurrent.Range("J" & r).Value
            wshDEB_Saisie.Range("K" & rowDEB).Value = wshDEB_Récurrent.Range("K" & r).Value
            wshDEB_Saisie.Range("L" & rowDEB).Value = wshDEB_Récurrent.Range("L" & r).Value
            wshDEB_Saisie.Range("M" & rowDEB).Value = wshDEB_Récurrent.Range("M" & r).Value
'            wshDEB_Saisie.Range("N" & rowJE).value = wshDEB_Récurrent.Range("I" & r).value
            wshDEB_Saisie.Range("Q" & rowDEBAuto).Value = wshDEB_Récurrent.Range("F" & r).Value
            totAmount = totAmount + wshDEB_Récurrent.Range("I" & r).Value
            If typeDEB = "" Then
                typeDEB = wshDEB_Récurrent.Range("C" & r).Value
            End If
            rowDEB = rowDEB + 1
        End If
    Next r
    wshDEB_Saisie.Range("F4").Value = typeDEB
    wshDEB_Saisie.Range("F6").Value = "[Auto]-" & DEBAutoDesc
    wshDEB_Saisie.Range("O6").Value = Format$(totAmount, "#,##0.00")
    wshDEB_Saisie.Range("O4").Select
    wshDEB_Saisie.Range("O4").Activate

    Application.EnableEvents = True

    Call Log_Record("modGL_EJ:Load_JEAuto_Into_JE", startTime)
    
End Sub

Sub Save_DEB_Recurrent(ll As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:Save_DEB_Recurrent", 0)
    
    Dim rowDEBLast As Long
    rowDEBLast = wshDEB_Saisie.Cells(wshDEB_Saisie.Rows.count, "E").End(xlUp).row  'Last Used Row in wshDEB_Saisie
    
    Call DEB_Recurrent_Add_Record_To_DB(rowDEBLast)
    Call DEB_Recurrent_Add_Record_Locally(rowDEBLast)
    
    Call Log_Record("modDEB_Saisie:Save_DEB_Recurrent", startTime)
    
End Sub

Sub DEB_Recurrent_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Recurrent_Add_Record_To_DB", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
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

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields(fDebRNoDebRec - 1).Value = nextDRNo
            rs.Fields(fDebRDate - 1).Value = wshDEB_Saisie.Range("O4").Value
            rs.Fields(fDebRType - 1).Value = wshDEB_Saisie.Range("F4").Value
            rs.Fields(fDebRBeneficiaire - 1).Value = wshDEB_Saisie.Range("J4").Value
            rs.Fields(fDebRReference - 1).Value = wshDEB_Saisie.Range("M6").Value
            rs.Fields(fDebRNoCompte - 1).Value = wshDEB_Saisie.Range("Q" & l).Value
            rs.Fields(fDebRCompte - 1).Value = wshDEB_Saisie.Range("E" & l).Value
            rs.Fields(fDebRCodeTaxe - 1).Value = wshDEB_Saisie.Range("H" & l).Value
            rs.Fields(fDebRTotal - 1).Value = wshDEB_Saisie.Range("I" & l).Value
            rs.Fields(fDebRTPS - 1).Value = wshDEB_Saisie.Range("J" & l).Value
            rs.Fields(fDebRTVQ - 1).Value = wshDEB_Saisie.Range("K" & l).Value
            rs.Fields(fDebRCréditTPS - 1).Value = wshDEB_Saisie.Range("L" & l).Value
            rs.Fields(fDebRCréditTVQ - 1).Value = wshDEB_Saisie.Range("M" & l).Value
            rs.Fields(fDebRTimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
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
    
    Call Log_Record("modDEB_Saisie:DEB_Recurrent_Add_Record_To_DB", startTime)

End Sub

Sub DEB_Recurrent_Add_Record_Locally(r As Long) 'Write records to local file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Recurrent_Add_Record_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim DEBRecNo As Long
    DEBRecNo = wshDEB_Saisie.Range("B2").Value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshDEB_Récurrent.Cells(wshDEB_Récurrent.Rows.count, "C").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Long
    For i = 9 To r
        wshDEB_Récurrent.Range("A" & rowToBeUsed).Value = DEBRecNo
        wshDEB_Récurrent.Range("B" & rowToBeUsed).Value = wshDEB_Saisie.Range("O4").Value
        wshDEB_Récurrent.Range("C" & rowToBeUsed).Value = wshDEB_Saisie.Range("F4").Value
        wshDEB_Récurrent.Range("D" & rowToBeUsed).Value = wshDEB_Saisie.Range("J4").Value
        wshDEB_Récurrent.Range("E" & rowToBeUsed).Value = wshDEB_Saisie.Range("M6").Value
        
        wshDEB_Récurrent.Range("F" & rowToBeUsed).Value = wshDEB_Saisie.Range("Q" & i).Value
        wshDEB_Récurrent.Range("G" & rowToBeUsed).Value = wshDEB_Saisie.Range("E" & i).Value
        wshDEB_Récurrent.Range("H" & rowToBeUsed).Value = wshDEB_Saisie.Range("H" & i).Value
        wshDEB_Récurrent.Range("I" & rowToBeUsed).Value = wshDEB_Saisie.Range("I" & i).Value
        wshDEB_Récurrent.Range("J" & rowToBeUsed).Value = wshDEB_Saisie.Range("J" & i).Value
        wshDEB_Récurrent.Range("K" & rowToBeUsed).Value = wshDEB_Saisie.Range("K" & i).Value
        wshDEB_Récurrent.Range("L" & rowToBeUsed).Value = wshDEB_Saisie.Range("L" & i).Value
        wshDEB_Récurrent.Range("M" & rowToBeUsed).Value = wshDEB_Saisie.Range("M" & i).Value
        wshDEB_Récurrent.Range("N" & rowToBeUsed).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call DEB_Recurrent_Build_Summary '2024-03-14 @ 07:40
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modDEB_Saisie:DEB_Recurrent_Add_Record_Locally", startTime)
    
End Sub

Sub DEB_Recurrent_Build_Summary()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Recurrent_Build_Summary", 0)
    
    'Build the summary at column K & L
    Dim lastUsedRow1 As Long
    lastUsedRow1 = wshDEB_Récurrent.Cells(wshDEB_Récurrent.Rows.count, "A").End(xlUp).row
    
    Dim lastUsedRow2 As Long
    lastUsedRow2 = wshDEB_Récurrent.Cells(wshDEB_Récurrent.Rows.count, "O").End(xlUp).row
    If lastUsedRow2 > 1 Then
        wshDEB_Récurrent.Range("P2:R" & lastUsedRow2).ClearContents
    End If
    
    With wshDEB_Récurrent
        Dim i As Long, k As Long, oldEntry As String
        k = 2
        For i = 2 To lastUsedRow1
            If .Range("A" & i).Value <> oldEntry Then
                .Range("P" & k).Value = "'" & Fn_Pad_A_String(.Range("A" & i).Value, " ", 5, "L")
                .Range("Q" & k).Value = .Range("D" & i).Value
                .Range("R" & k).Value = .Range("B" & i).Value
                oldEntry = .Range("A" & i).Value
                k = k + 1
            End If
        Next i
    End With

    Call Log_Record("modDEB_Saisie:DEB_Recurrent_Build_Summary", startTime)

End Sub

Public Sub DEB_Saisie_Clear_All_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Saisie_Clear_All_Cells", 0)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    With wshDEB_Saisie
        .Range("F4:H4, J4:M4, F6:J6, M6, O6, E9:O23, Q9:Q23").ClearContents
        .Range("O4").Value = Format$(Date, wshAdmin.Range("B1").Value)
        .ckbRecurrente = False
    End With
    
    'Toutes les cellules sont sans surbrillance (élimine le vert pâle)
    With wshDEB_Saisie.Range("F4:H4, J4:M4, F6:J6, M6, O6, E9:O23, I26, L26:O26").Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'Protection de la feuille, seules les cellules non-verrouillées peuvent être sélectionnées
    With wshDEB_Saisie
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Call Log_Record("modDEB_Saisie:DEB_Saisie_Clear_All_Cells", startTime)

End Sub

Sub shp_DEB_Back_To_Menu_Click()

    Call DEBOURS_Back_To_Menu

End Sub

Sub DEBOURS_Back_To_Menu()
    
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
        'GST calculation
        If taxCode = "TPS/TVQ" Or taxCode = "REP" Then
            gst = Round(total / (1 + gstRate + pstRate) * gstRate, 2)
        Else
            gst = 0
        End If
        
        'PST calculation
        If taxCode = "TPS/TVQ" Or taxCode = "REP" Then
            pst = Round(total / (1 + gstRate + pstRate) * pstRate, 2)
        Else
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
        
        'Net amount (Expense) = Total - gstCredit - pstCredit
        netAmount = total - gstCredit - pstCredit
        Exit Sub
    End If
    
    If netAmount <> 0 Then 'Calculate the taxes from the net amount
        'gst calculation
        If taxCode = "TPS/TVQ" Or taxCode = "REP" Then
            gst = Round(netAmount * gstRate, 2)
        Else
            gst = 0
        End If
        
        'PST calculation
        If taxCode = "TPS/TVQ" Or taxCode = "REP" Then
            pst = Round(netAmount * pstRate, 2)
        Else
            pst = 0
        End If
        
        If taxCode = "REP" Then
            gstCredit = Round(gst / 2, 2)
            pstCredit = Round(pst / 2, 2)
        Else
            gstCredit = gst
            pstCredit = pst
        End If
        
        total = netAmount + gstCredit + pstCredit
        
    End If
    
End Sub

