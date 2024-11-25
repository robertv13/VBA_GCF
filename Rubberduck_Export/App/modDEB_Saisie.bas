Attribute VB_Name = "modDEB_Saisie"
Option Explicit

Sub DEB_Saisie_Update()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Saisie_Update", 0)
    
    'Remove highlight from last cell
    If wshDEB_Saisie.Range("B4").value <> "" Then
        wshDEB_Saisie.Range(wshDEB_Saisie.Range("B4").value).Interior.Color = xlNone
    End If
    
    'Date is not valid OR the transaction does not balance
    If Fn_Is_Date_Valide(wshDEB_Saisie.Range("O4").value) = False Or _
        Fn_Is_Debours_Balance = False Then
            Exit Sub
    End If
    
    'Is every line of the transaction well entered ?
    Dim rowDebSaisie As Long
    rowDebSaisie = wshDEB_Saisie.Range("E23").End(xlUp).row  'Last Used Row in wshDEB_Saisie
    If Fn_Is_Deb_Saisie_Valid(rowDebSaisie) = False Then Exit Sub
    
    'Get the Fourn_ID
    wshDEB_Saisie.Range("B5").value = Fn_GetID_From_Fourn_Name(wshDEB_Saisie.Range("J4").value)

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
    CurrentDeboursNo = wshDEB_Saisie.Range("B1").value
    
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
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "DEB_Trans$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(No_Entrée) AS MaxDebTransNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastDebTrans As Long
    If IsNull(rs.Fields("MaxDebTransNo").value) Then
        'Handle empty table (assign a default value, e.g., 0)
        lastDebTrans = 0
    Else
        lastDebTrans = rs.Fields("MaxDebTransNo").value
    End If
    
    'Calculate the new JE number
    Dim currDebTransNo As Long
    currDebTransNo = lastDebTrans + 1
    Application.EnableEvents = False
    wshDEB_Saisie.Range("B1").value = currDebTransNo
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
            rs.Fields("No_Entrée").value = currDebTransNo
            rs.Fields("Date").value = wshDEB_Saisie.Range("O4").value
            rs.Fields("Type").value = wshDEB_Saisie.Range("F4").value
            rs.Fields("Beneficiaire").value = wshDEB_Saisie.Range("J4").value
            rs.Fields("FournID").value = wshDEB_Saisie.Range("B5").value
            rs.Fields("Description").value = wshDEB_Saisie.Range("F6").value
            rs.Fields("Reference").value = wshDEB_Saisie.Range("M6").value
            rs.Fields("Total").value = wshDEB_Saisie.Range("O6").value
            rs.Fields("No_Compte").value = wshDEB_Saisie.Range("Q" & l).value
            rs.Fields("Compte").value = wshDEB_Saisie.Range("E" & l).value
            rs.Fields("CodeTaxe").value = wshDEB_Saisie.Range("H" & l).value
            rs.Fields("Total").value = CDbl(wshDEB_Saisie.Range("I" & l).value)
            rs.Fields("TPS").value = CDbl(wshDEB_Saisie.Range("J" & l).value)
            rs.Fields("TVQ").value = CDbl(wshDEB_Saisie.Range("K" & l).value)
            rs.Fields("Crédit_TPS").value = CDbl(wshDEB_Saisie.Range("L" & l).value)
            rs.Fields("Crédit_TVQ").value = CDbl(wshDEB_Saisie.Range("M" & l).value)
            'Montant de dépense (Total - creditTPS - creditTVQ)
            rs.Fields("Dépense").value = CDbl(wshDEB_Saisie.Range("I" & l).value _
                                              - wshDEB_Saisie.Range("L" & l).value _
                                              - wshDEB_Saisie.Range("M" & l).value)
            rs.Fields("AutreRemarque").value = ""
            rs.Fields("TimeStamp").value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rs.update
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
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentDebTransNo As Long
    currentDebTransNo = wshDEB_Saisie.Range("B1").value
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshDEB_Trans.Cells(wshDEB_Trans.Rows.count, "A").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Long
    For i = 9 To r
        wshDEB_Trans.Range("A" & rowToBeUsed).value = currentDebTransNo
        wshDEB_Trans.Range("B" & rowToBeUsed).value = wshDEB_Saisie.Range("O4").value
        wshDEB_Trans.Range("C" & rowToBeUsed).value = wshDEB_Saisie.Range("F4").value
        wshDEB_Trans.Range("D" & rowToBeUsed).value = wshDEB_Saisie.Range("J4").value
        wshDEB_Trans.Range("E" & rowToBeUsed).value = wshDEB_Saisie.Range("B5").value
        wshDEB_Trans.Range("F" & rowToBeUsed).value = wshDEB_Saisie.Range("F6").value
        wshDEB_Trans.Range("G" & rowToBeUsed).value = wshDEB_Saisie.Range("M6").value
        wshDEB_Trans.Range("H" & rowToBeUsed).value = wshDEB_Saisie.Range("Q" & i).value
        wshDEB_Trans.Range("I" & rowToBeUsed).value = wshDEB_Saisie.Range("E" & i).value
        wshDEB_Trans.Range("J" & rowToBeUsed).value = wshDEB_Saisie.Range("H" & i).value
        wshDEB_Trans.Range("K" & rowToBeUsed).value = wshDEB_Saisie.Range("I" & i).value
        wshDEB_Trans.Range("L" & rowToBeUsed).value = wshDEB_Saisie.Range("J" & i).value
        wshDEB_Trans.Range("M" & rowToBeUsed).value = wshDEB_Saisie.Range("K" & i).value
        wshDEB_Trans.Range("N" & rowToBeUsed).value = wshDEB_Saisie.Range("L" & i).value
        wshDEB_Trans.Range("O" & rowToBeUsed).value = wshDEB_Saisie.Range("M" & i).value
        '$ dépense = Total - creditTPS - creditTVQ
        wshDEB_Trans.Range("P" & rowToBeUsed).value = wshDEB_Saisie.Range("I" & i).value _
                                                      - wshDEB_Saisie.Range("L" & i).value _
                                                      - wshDEB_Saisie.Range("M" & i).value
        wshDEB_Trans.Range("Q" & rowToBeUsed).value = ""
        wshDEB_Trans.Range("R" & rowToBeUsed).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_Locally", startTime)

    Application.ScreenUpdating = True

End Sub

Sub DEB_Saisie_GL_Posting_Preparation() '2024-06-05 @ 18:28

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Saisie_GL_Posting_Preparation", 0)

    Dim montant As Double, dateDebours As Date
    Dim descGL_Trans As String, source As String, deboursType As String
    Dim GL_TransNo As Long
    
    dateDebours = wshDEB_Saisie.Range("O4").value
    deboursType = wshDEB_Saisie.Range("F4").value
    descGL_Trans = deboursType & " - " & wshDEB_Saisie.Range("F6").value
    If Trim(wshDEB_Saisie.Range("M6").value) <> "" Then
        descGL_Trans = descGL_Trans & " [" & wshDEB_Saisie.Range("M6").value & "]"
    End If
    source = "DÉBOURSÉ:" & Format$(wshDEB_Saisie.Range("B1").value, "00000")
    
    Dim MyArray() As String
    ReDim MyArray(1 To 16, 1 To 4)
    
    'Based on Disbursement type, the CREDIT account will be different
    'Disbursement Total (wshDEB_Saisie.Range("O6"))
    montant = wshDEB_Saisie.Range("O6").value
    
    Dim GLNo_Credit As String
    
    Select Case deboursType
        Case "Chèque", "Virement", "Paiement pré-autorisé"
            MyArray(1, 1) = "1000"
            MyArray(1, 2) = "Encaisse"
        Case "Carte de crédit"
            MyArray(1, 1) = "2010"
            MyArray(1, 2) = "Carte de crédit"
        Case "Avances avec Guillaume Charron"
            MyArray(1, 1) = "2200"
            MyArray(1, 2) = "Avances avec Guillaume Charron"
        Case "Avances avec 9249-3626 Québec inc."
            MyArray(1, 1) = "2210"
            MyArray(1, 2) = "Avances avec 9249-3626 Québec inc."
        Case "Avances avec 9333-4829 Québec inc."
            MyArray(1, 1) = "2220"
            MyArray(1, 2) = "Avances avec 9333-4829 Québec inc."
        Case "Autre"
            MyArray(1, 1) = "1000"
            MyArray(1, 2) = "Encaisse"
        Case Else
            MyArray(1, 1) = "1000"
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
        MyArray(arrRow, 1) = wshDEB_Saisie.Range("Q" & l).value
        MyArray(arrRow, 2) = wshDEB_Saisie.Range("E" & l).value
        MyArray(arrRow, 3) = wshDEB_Saisie.Range("N" & l).value
        MyArray(arrRow, 4) = ""
        arrRow = arrRow + 1
        
        If wshDEB_Saisie.Range("L" & l).value <> 0 Then
            MyArray(arrRow, 1) = "1200"
            MyArray(arrRow, 2) = "TPS payées"
            MyArray(arrRow, 3) = wshDEB_Saisie.Range("L" & l).value
            MyArray(arrRow, 4) = ""
            arrRow = arrRow + 1
        End If

        If wshDEB_Saisie.Range("M" & l).value <> 0 Then
            MyArray(arrRow, 1) = "1201"
            MyArray(arrRow, 2) = "TVQ payées"
            MyArray(arrRow, 3) = wshDEB_Saisie.Range("M" & l).value
            MyArray(arrRow, 4) = ""
            arrRow = arrRow + 1
        End If
    Next l
    
    Dim glEntryNo As Long
    Call GL_Posting_To_DB(dateDebours, descGL_Trans, source, MyArray, glEntryNo)
    
    Call GL_Posting_Locally(dateDebours, descGL_Trans, source, MyArray, glEntryNo)
    
    Call Log_Record("modDEB_Saisie:DEB_Saisie_GL_Posting_Preparation", startTime)

End Sub

Sub Load_DEB_Auto_Into_JE(DEBAutoDesc As String, NoDEBAuto As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:Load_DEB_Auto_Into_JE", 0)
    
    'On copie l'écriture automatique vers wshDEB_Saisie
    Dim rowDEBAuto, rowDEB As Long
    rowDEBAuto = wshDEB_Recurrent.Cells(wshDEB_Recurrent.Rows.count, "C").End(xlUp).row  'Last Row used in wshDEB_Recuurent
    
    Call DEB_Saisie_Clear_All_Cells
    
    rowDEB = 9
    
    Application.EnableEvents = False
    Dim r As Long, totAmount As Currency, typeDEB As String
    For r = 2 To rowDEBAuto
        If wshDEB_Recurrent.Range("A" & r).value = NoDEBAuto And wshDEB_Recurrent.Range("F" & r).value <> "" Then
            wshDEB_Saisie.Range("E" & rowDEB).value = wshDEB_Recurrent.Range("G" & r).value
            wshDEB_Saisie.Range("H" & rowDEB).value = wshDEB_Recurrent.Range("H" & r).value
            wshDEB_Saisie.Range("I" & rowDEB).value = wshDEB_Recurrent.Range("I" & r).value
            wshDEB_Saisie.Range("J" & rowDEB).value = wshDEB_Recurrent.Range("J" & r).value
            wshDEB_Saisie.Range("K" & rowDEB).value = wshDEB_Recurrent.Range("K" & r).value
            wshDEB_Saisie.Range("L" & rowDEB).value = wshDEB_Recurrent.Range("L" & r).value
            wshDEB_Saisie.Range("M" & rowDEB).value = wshDEB_Recurrent.Range("M" & r).value
'            wshDEB_Saisie.Range("N" & rowJE).value = wshDEB_Recurrent.Range("I" & r).value
            wshDEB_Saisie.Range("Q" & rowDEBAuto).value = wshDEB_Recurrent.Range("F" & r).value
            totAmount = totAmount + wshDEB_Recurrent.Range("I" & r).value
            If typeDEB = "" Then
                typeDEB = wshDEB_Recurrent.Range("C" & r).value
            End If
            rowDEB = rowDEB + 1
        End If
    Next r
    wshDEB_Saisie.Range("F4").value = typeDEB
    wshDEB_Saisie.Range("F6").value = "[Auto]-" & DEBAutoDesc
    wshDEB_Saisie.Range("O6").value = Format$(totAmount, "#,##0.00")
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
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "DEB_Recurrent$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxDebRecNo As Long
    strSQL = "SELECT MAX(No_Deb_Rec) AS MaxDebRecNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastDR As Long, nextDRNo As Long
    If IsNull(rs.Fields("MaxDebRecNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastDR = 0
    Else
        lastDR = rs.Fields("MaxDebRecNo").value
    End If
    
    'Calculate the new ID
    nextDRNo = lastDR + 1
    wshDEB_Saisie.Range("B2").value = nextDRNo

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("No_Deb_Rec").value = nextDRNo
            rs.Fields("Date").value = wshDEB_Saisie.Range("O4").value
            rs.Fields("Type").value = wshDEB_Saisie.Range("F4").value
            rs.Fields("Beneficiaire").value = wshDEB_Saisie.Range("F6").value
            rs.Fields("Reference").value = wshDEB_Saisie.Range("M6").value
            
            rs.Fields("No_Compte").value = wshDEB_Saisie.Range("Q" & l).value
            rs.Fields("Compte").value = wshDEB_Saisie.Range("E" & l).value
            rs.Fields("CodeTaxe").value = wshDEB_Saisie.Range("H" & l).value
            rs.Fields("Total").value = wshDEB_Saisie.Range("I" & l).value
            rs.Fields("TPS").value = wshDEB_Saisie.Range("J" & l).value
            rs.Fields("TVQ").value = wshDEB_Saisie.Range("K" & l).value
            rs.Fields("Crédit_TPS").value = wshDEB_Saisie.Range("L" & l).value
            rs.Fields("Crédit_TVQ").value = wshDEB_Saisie.Range("M" & l).value
        rs.update
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
    DEBRecNo = wshDEB_Saisie.Range("B2").value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshDEB_Recurrent.Cells(wshDEB_Recurrent.Rows.count, "C").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Long
    For i = 9 To r
        wshDEB_Recurrent.Range("A" & rowToBeUsed).value = DEBRecNo
        wshDEB_Recurrent.Range("B" & rowToBeUsed).value = wshDEB_Saisie.Range("O4").value
        wshDEB_Recurrent.Range("C" & rowToBeUsed).value = wshDEB_Saisie.Range("F4").value
        wshDEB_Recurrent.Range("D" & rowToBeUsed).value = wshDEB_Saisie.Range("F6").value
        wshDEB_Recurrent.Range("E" & rowToBeUsed).value = wshDEB_Saisie.Range("M6").value
        
        wshDEB_Recurrent.Range("F" & rowToBeUsed).value = wshDEB_Saisie.Range("Q" & i).value
        wshDEB_Recurrent.Range("G" & rowToBeUsed).value = wshDEB_Saisie.Range("E" & i).value
        wshDEB_Recurrent.Range("H" & rowToBeUsed).value = wshDEB_Saisie.Range("H" & i).value
        wshDEB_Recurrent.Range("I" & rowToBeUsed).value = wshDEB_Saisie.Range("I" & i).value
        wshDEB_Recurrent.Range("J" & rowToBeUsed).value = wshDEB_Saisie.Range("J" & i).value
        wshDEB_Recurrent.Range("K" & rowToBeUsed).value = wshDEB_Saisie.Range("K" & i).value
        wshDEB_Recurrent.Range("L" & rowToBeUsed).value = wshDEB_Saisie.Range("L" & i).value
        wshDEB_Recurrent.Range("M" & rowToBeUsed).value = wshDEB_Saisie.Range("M" & i).value
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
    lastUsedRow1 = wshDEB_Recurrent.Cells(wshDEB_Recurrent.Rows.count, "A").End(xlUp).row
    
    Dim lastUsedRow2 As Long
    lastUsedRow2 = wshDEB_Recurrent.Cells(wshDEB_Recurrent.Rows.count, "O").End(xlUp).row
    If lastUsedRow2 > 1 Then
        wshDEB_Recurrent.Range("O2:Q" & lastUsedRow2).ClearContents
    End If
    
    With wshDEB_Recurrent
        Dim i As Long, k As Long, oldEntry As String
        k = 2
        For i = 2 To lastUsedRow1
            If .Range("A" & i).value <> oldEntry Then
                .Range("O" & k).value = "'" & Fn_Pad_A_String(.Range("A" & i).value, " ", 5, "L")
                .Range("P" & k).value = .Range("D" & i).value
                .Range("Q" & k).value = .Range("B" & i).value
                oldEntry = .Range("A" & i).value
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
        .Range("O4").value = Format$(Now(), wshAdmin.Range("B1").value)
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

Sub DEBOURS_Back_To_Menu()
    
    wshDEB_Saisie.Visible = xlSheetHidden
    
    Application.ScreenUpdating = False
    
    wshMenuDEB.Activate
    wshMenuDEB.Range("A1").Select
    
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

