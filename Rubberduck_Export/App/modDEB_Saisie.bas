Attribute VB_Name = "modDEB_Saisie"
Option Explicit

Sub DEB_Saisie_Update()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:DEB_Saisie_Update()")
    
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
    wshDEB_Saisie.Range("B5").value = Fn_FournID_From_Fourn_Name(wshDEB_Saisie.Range("F6").value)

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
    
    wshDEB_Saisie.Range("F4").Select
        
    Call Output_Timer_Results("modDEB_Saisie:DEB_Saisie_Update()", timerStart)
        
End Sub

Sub DEB_Trans_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:DEB_Trans_Add_Record_To_DB()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = rootPath & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "DEB_Trans"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(No_Entrée) AS MaxDebTransNo FROM [" & destinationTab & "$]"

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
    
    'Build formula
    Dim formula As String
    formula = "=ROW()"

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    'Read all line from Journal Entry
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("No_Entrée").value = currDebTransNo
            rs.Fields("Date").value = CDate(wshDEB_Saisie.Range("O4").value)
            rs.Fields("Type").value = wshDEB_Saisie.Range("F4").value
            rs.Fields("Beneficiaire").value = wshDEB_Saisie.Range("F6").value
            rs.Fields("FournID").value = wshDEB_Saisie.Range("B5").value
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
            rs.Fields("AutreRemarque").value = ""
'            rs.Fields("TimeStamp").value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
            rs.Fields("TimeStamp").value = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
            Debug.Print "DEB_Trans - " & CDate(Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        rs.update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Output_Timer_Results("modDEB_Saisie:DEB_Trans_Add_Record_To_DB()", timerStart)

End Sub

Sub DEB_Trans_Add_Record_Locally(r As Long) 'Write records locally
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:DEB_Trans_Add_Record_Locally()")
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentDebTransNo As Long
    currentDebTransNo = wshDEB_Saisie.Range("B1").value
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshDEB_Trans.Range("A99999").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Long
    For i = 9 To r
        wshDEB_Trans.Range("A" & rowToBeUsed).value = currentDebTransNo
        wshDEB_Trans.Range("B" & rowToBeUsed).value = CDate(wshDEB_Saisie.Range("O4").value)
        wshDEB_Trans.Range("C" & rowToBeUsed).value = wshDEB_Saisie.Range("F4").value
        wshDEB_Trans.Range("D" & rowToBeUsed).value = wshDEB_Saisie.Range("F6").value
        wshDEB_Trans.Range("E" & rowToBeUsed).value = wshDEB_Saisie.Range("B5").value
        wshDEB_Trans.Range("F" & rowToBeUsed).value = wshDEB_Saisie.Range("M6").value
        wshDEB_Trans.Range("G" & rowToBeUsed).value = wshDEB_Saisie.Range("Q" & i).value
        wshDEB_Trans.Range("H" & rowToBeUsed).value = wshDEB_Saisie.Range("E" & i).value
        wshDEB_Trans.Range("I" & rowToBeUsed).value = wshDEB_Saisie.Range("H" & i).value
        wshDEB_Trans.Range("J" & rowToBeUsed).value = wshDEB_Saisie.Range("I" & i).value
        wshDEB_Trans.Range("K" & rowToBeUsed).value = wshDEB_Saisie.Range("J" & i).value
        wshDEB_Trans.Range("L" & rowToBeUsed).value = wshDEB_Saisie.Range("K" & i).value
        wshDEB_Trans.Range("M" & rowToBeUsed).value = wshDEB_Saisie.Range("L" & i).value
        wshDEB_Trans.Range("N" & rowToBeUsed).value = wshDEB_Saisie.Range("M" & i).value
        wshDEB_Trans.Range("O" & rowToBeUsed).value = ""
        wshDEB_Trans.Range("P" & rowToBeUsed).value = Format$(Now(), "mm/dd/yyyy hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call Output_Timer_Results("modDEB_Saisie:DEB_Trans_Add_Record_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub DEB_Saisie_GL_Posting_Preparation() '2024-06-05 @ 18:28

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:DEB_Saisie_GL_Posting_Preparation()")

    Dim montant As Double, dateDebours As Date
    Dim descGL_Trans As String, source As String, deboursType As String
    Dim GL_TransNo As Long
    
    dateDebours = wshDEB_Saisie.Range("O4").value
    deboursType = wshDEB_Saisie.Range("F4").value
    descGL_Trans = deboursType & " - " & wshDEB_Saisie.Range("F6").value
    If Trim(wshDEB_Saisie.Range("M6").value) <> "" Then
        descGL_Trans = descGL_Trans & " [" & wshDEB_Saisie.Range("M6").value & "]"
    End If
    source = "DÉBOURS-" & Format$(wshDEB_Saisie.Range("B1").value, "000000")
    
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
        Case "VISA", "MCARD", "AMEX"
            MyArray(1, 1) = "2010"
            MyArray(1, 2) = "Carte de crédit"
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
    lastUsedRow = wshDEB_Saisie.Range("E99").End(xlUp).row

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
    
    Call GL_Posting_To_DB(dateDebours, descGL_Trans, source, MyArray)
    GL_TransNo = wshAdmin.Range("B9").value
    Call GL_Posting_Locally(dateDebours, descGL_Trans, source, GL_TransNo, MyArray)
    
    Call Output_Timer_Results("modDEB_Saisie:DEB_Saisie_GL_Posting_Preparation()", timerStart)

End Sub

Sub Load_DEB_Auto_Into_JE(DEBAutoDesc As String, NoDEBAuto As Long)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:Load_DEB_Auto_Into_JE()")
    
    'On copie l'écriture automatique vers wshDEB_Saisie
    Dim rowDEBAuto, rowDEB As Long
    rowDEBAuto = wshDEB_Recurrent.Range("C99999").End(xlUp).row  'Last Row used in wshDEB_Recuurent
    
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

    Call Output_Timer_Results("modGL_EJ:Load_JEAuto_Into_JE()", timerStart)
    
End Sub

Sub Save_DEB_Recurrent(ll As Long)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:Save_DEB_Recurrent()")
    
    Dim rowDEBLast As Long
    rowDEBLast = wshDEB_Saisie.Range("E99").End(xlUp).row  'Last Used Row in wshDEB_Saisie
    
    Call DEB_Recurrent_Add_Record_To_DB(rowDEBLast)
    Call DEB_Recurrent_Add_Record_Locally(rowDEBLast)
    
    Call Output_Timer_Results("modDEB_Saisie:Save_DEB_Recurrent()", timerStart)
    
End Sub

Sub DEB_Recurrent_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:DEB_Recurrent_Add_Record_To_DB()")

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = rootPath & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "DEB_Recurrent"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxDebRecNo As Long
    strSQL = "SELECT MAX(No_Deb_Rec) AS MaxDebRecNo FROM [" & destinationTab & "$]"

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
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("No_Deb_Rec").value = nextDRNo
            rs.Fields("Date").value = CDate(wshDEB_Saisie.Range("O4").value)
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

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Output_Timer_Results("modDEB_Saisie:DEB_Recurrent_Add_Record_To_DB()", timerStart)

End Sub

Sub DEB_Recurrent_Add_Record_Locally(r As Long) 'Write records to local file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:DEB_Recurrent_Add_Record_Locally()")
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim DEBRecNo As Long
    DEBRecNo = wshDEB_Saisie.Range("B2").value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshDEB_Recurrent.Range("C999").End(xlUp).row
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
    
    Call Output_Timer_Results("modDEB_Saisie:DEB_Recurrent_Add_Record_Locally()", timerStart)
    
End Sub

Sub DEB_Recurrent_Build_Summary()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:DEB_Recurrent_Build_Summary()")
    
    'Build the summary at column K & L
    Dim lastUsedRow1 As Long
    lastUsedRow1 = wshDEB_Recurrent.Range("A9999").End(xlUp).row
    
    Dim lastUsedRow2 As Long
    lastUsedRow2 = wshDEB_Recurrent.Range("O999").End(xlUp).row
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

    Call Output_Timer_Results("modDEB_Saisie:DEB_Recurrent_Build_Summary()", timerStart)

End Sub

Public Sub DEB_Saisie_Clear_All_Cells()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modDEB_Saisie:DEB_Saisie_Clear_All_Cells()")

    'Vide les cellules
    Application.EnableEvents = False
    With wshDEB_Saisie
        .Range("F4:H4, F6:K6, M6, O6, E9:O23, Q9:Q23").ClearContents
        .Range("O4").value = Format$(Now(), "mm/dd/yyyy")
        .ckbRecurrente = False
    End With
    Application.EnableEvents = True
    
    Call Output_Timer_Results("modDEB_Saisie:DEB_Saisie_Clear_All_Cells()", timerStart)

End Sub

Sub DEBOURS_Back_To_Menu()
    
    wshDEB_Saisie.Visible = xlSheetHidden
    
    wshMenuDEB.Activate
    wshMenuDEB.Range("A1").Select
    
End Sub


