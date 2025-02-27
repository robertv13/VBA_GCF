Attribute VB_Name = "modGL_JE"
Option Explicit

Sub JE_Update()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_JE:JE_Update()")
    
    If Fn_Is_Date_Valide(wshGL_EJ.Range("K4").Value) = False Then Exit Sub
    
    If Fn_Is_Ecriture_Balance = False Then Exit Sub
    
    Dim rowEJLast As Long
    rowEJLast = wshGL_EJ.Range("E23").End(xlUp).row  'Last Used Row in wshGL_EJ
    If Fn_Is_JE_Valid(rowEJLast) = False Then Exit Sub
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    Call GL_Trans_Add_Record_To_DB(rowEJLast)
    Call GL_Trans_Add_Record_Locally(rowEJLast)
    
    If wshGL_EJ.ckbRecurrente = True Then
        Call Save_EJ_Recurrente(rowEJLast)
    End If
    
    'Save Current JE number
    Dim strCurrentJE As String
    strCurrentJE = wshGL_EJ.Range("B1").Value
    
    'Increment Next JE number
    wshGL_EJ.Range("B1").Value = wshGL_EJ.Range("B1").Value + 1
        
    Call wshGL_EJ_Clear_All_Cells
        
    With wshGL_EJ
        .Activate
        .Range("F4").Select
        .Range("F4").Activate
    End With
    
    MsgBox "L'écriture numéro '" & strCurrentJE & "' a été reporté avec succès"
    
    Call Output_Timer_Results("modGL_JE:JE_Update()", timerStart)
    
End Sub

Sub Save_EJ_Recurrente(ll As Long)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_JE:Save_EJ_Recurrente()")
    
    Dim rowEJLast As Long
    rowEJLast = wshGL_EJ.Range("E99").End(xlUp).row  'Last Used Row in wshGL_EJ
    
    Call GL_EJ_Auto_Add_Record_To_DB(ll)
    Call GL_EJ_Auto_Add_Record_Locally(ll)
    
    Call Output_Timer_Results("modGL_JE:Save_EJ_Recurrente()", timerStart)
    
End Sub

Sub Load_JEAuto_Into_JE(EJAutoDesc As String, NoEJAuto As Long)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_JE:Load_JEAuto_Into_JE()")
    
    'On copie l'E/J automatique vers wshEJ
    Dim rowJEAuto, rowJE As Long
    rowJEAuto = wshGL_EJ_Recurrente.Range("C99999").End(xlUp).row  'Last Row used in wshGL_EJRecuurente
    
    Call wshGL_EJ_Clear_All_Cells
    rowJE = 9
    
    Dim r As Long
    For r = 2 To rowJEAuto
        If wshGL_EJ_Recurrente.Range("C" & r).Value = NoEJAuto And wshGL_EJ_Recurrente.Range("E" & r).Value <> "" Then
            wshGL_EJ.Range("E" & rowJE).Value = wshGL_EJ_Recurrente.Range("F" & r).Value
            wshGL_EJ.Range("H" & rowJE).Value = wshGL_EJ_Recurrente.Range("G" & r).Value
            wshGL_EJ.Range("I" & rowJE).Value = wshGL_EJ_Recurrente.Range("H" & r).Value
            wshGL_EJ.Range("J" & rowJE).Value = wshGL_EJ_Recurrente.Range("I" & r).Value
            wshGL_EJ.Range("L" & rowJE).Value = wshGL_EJ_Recurrente.Range("E" & r).Value
            rowJE = rowJE + 1
        End If
    Next r
    wshGL_EJ.Range("F6").Value = "[Auto]-" & EJAutoDesc
    wshGL_EJ.Range("K4").Activate

    Call Output_Timer_Results("modGL_JE:Load_JEAuto_Into_JE()", timerStart)
    
End Sub

Sub wshGL_EJ_Clear_All_Cells()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_JE:wshGL_EJ_Clear_All_Cells()")
    
    'Efface toutes les cellules de la feuille
    Application.EnableEvents = False
    ActiveSheet.Unprotect
    With wshGL_EJ
        .Range("F4,F6:K6").Clearcontents
        .Range("E9:G23,H9:H23,I9:I23,J9:L23").Clearcontents
        .ckbRecurrente = False
        Application.EnableEvents = True
        wshGL_EJ.Activate
        wshGL_EJ.Range("F4").Select
    End With
    ActiveSheet.Protect UserInterfaceOnly:=True
    
    Call Output_Timer_Results("modGL_JE:wshGL_EJ_Clear_All_Cells()", timerStart)

End Sub

Sub GL_EJ_Auto_Build_Summary()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_JE:GL_EJ_Auto_Build_Summary()")
    
    'Build the summary at column K & L
    Dim lastUsedRow1 As Long
    lastUsedRow1 = wshGL_EJ_Recurrente.Range("C999").End(xlUp).row
    
    Dim lastUsedRow2 As Long
    lastUsedRow2 = wshGL_EJ_Recurrente.Range("K999").End(xlUp).row
    If lastUsedRow2 > 1 Then
        wshGL_EJ_Recurrente.Range("K2:L" & lastUsedRow2).Clearcontents
    End If
    
    With wshGL_EJ_Recurrente
        Dim i As Integer, k As Integer, oldEntry As String
        k = 2
        For i = 2 To lastUsedRow1
            If .Range("D" & i).Value <> oldEntry Then
                .Range("K" & k).Value = .Range("D" & i).Value
                .Range("L" & k).Value = "'" & Fn_Pad_A_String(.Range("C" & i).Value, " ", 5, "L")
                oldEntry = .Range("D" & i).Value
                k = k + 1
            End If
        Next i
    End With

    Call Output_Timer_Results("modGL_JE:GL_EJ_Auto_Build_Summary()", timerStart)

End Sub

Sub GL_Trans_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_JE:GL_Trans_Add_Record_To_DB()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "GL_Trans"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(No_Entrée) AS MaxEJNo FROM [" & destinationTab & "$]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim maxEJNo As Long, lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").Value
    End If
    
    'Calculate the new JE number
    Dim nextJENo As Long
    nextJENo = lastJE + 1
    wshGL_EJ.Range("B1").Value = nextJENo
    
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
            rs.Fields("No_Entrée").Value = nextJENo
            rs.Fields("Date").Value = CDate(wshGL_EJ.Range("K4").Value)
            rs.Fields("Description").Value = wshGL_EJ.Range("F6").Value
            rs.Fields("Source").Value = wshGL_EJ.Range("F4").Value
            rs.Fields("No_Compte").Value = wshGL_EJ.Range("L" & l).Value
            rs.Fields("Compte").Value = wshGL_EJ.Range("E" & l).Value
            rs.Fields("Débit").Value = wshGL_EJ.Range("H" & l).Value
            rs.Fields("Crédit").Value = wshGL_EJ.Range("I" & l).Value
            rs.Fields("AutreRemarque").Value = wshGL_EJ.Range("J" & l).Value
            rs.Fields("TimeStamp").Value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        rs.update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("modGL_JE:GL_Trans_Add_Record_To_DB()", timerStart)

End Sub

Sub GL_Trans_Add_Record_Locally(r As Long) 'Write records locally
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_JE:GL_Trans_Add_Record_Locally()")
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshGL_EJ.Range("B1").Value
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_Trans.Range("A99999").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Integer
    For i = 9 To r
        wshGL_Trans.Range("A" & rowToBeUsed).Value = JENo
        wshGL_Trans.Range("B" & rowToBeUsed).Value = CDate(wshGL_EJ.Range("K4").Value)
        wshGL_Trans.Range("C" & rowToBeUsed).Value = wshGL_EJ.Range("F6").Value
        wshGL_Trans.Range("D" & rowToBeUsed).Value = wshGL_EJ.Range("F4").Value
        wshGL_Trans.Range("E" & rowToBeUsed).Value = wshGL_EJ.Range("L" & i).Value
        wshGL_Trans.Range("F" & rowToBeUsed).Value = wshGL_EJ.Range("E" & i).Value
        If wshGL_EJ.Range("H" & i).Value <> "" Then
            wshGL_Trans.Range("G" & rowToBeUsed).Value = wshGL_EJ.Range("H" & i).Value
        End If
        If wshGL_EJ.Range("I" & i).Value <> "" Then
            wshGL_Trans.Range("H" & rowToBeUsed).Value = wshGL_EJ.Range("I" & i).Value
        End If
        wshGL_Trans.Range("I" & rowToBeUsed).Value = wshGL_EJ.Range("J" & i).Value
        wshGL_Trans.Range("J" & rowToBeUsed).Value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("modGL_JE:GL_Trans_Add_Record_Locally()", timerStart)

End Sub

Sub GL_EJ_Auto_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_JE:GL_EJ_Auto_Add_Record_To_DB()")

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "GL_EJ_Auto"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJANo As Long
    strSQL = "SELECT MAX(No_EJA) AS MaxEJANo FROM [" & destinationTab & "$]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastEJA As Long, nextEJANo As Long
    If IsNull(rs.Fields("MaxEJANo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastEJA = 1
    Else
        lastEJA = rs.Fields("MaxEJANo").Value
    End If
    
    'Calculate the new ID
    nextEJANo = lastEJA + 1
    wshGL_EJ_Recurrente.Range("B2").Value = nextEJANo

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("No_EJA").Value = nextEJANo
            rs.Fields("Description").Value = wshGL_EJ.Range("F6").Value
            rs.Fields("No_Compte").Value = wshGL_EJ.Range("L" & l).Value
            rs.Fields("Compte").Value = wshGL_EJ.Range("E" & l).Value
            rs.Fields("Débit").Value = wshGL_EJ.Range("H" & l).Value
            rs.Fields("Crédit").Value = wshGL_EJ.Range("I" & l).Value
            rs.Fields("AutreRemarque").Value = wshGL_EJ.Range("J" & l).Value
        rs.update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("modGL_JE:GL_EJ_Auto_Add_Record_To_DB()", timerStart)

End Sub

Sub GL_EJ_Auto_Add_Record_Locally(r As Long) 'Write records to local file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_JE:GL_EJ_Auto_Add_Record_Locally()")
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshGL_EJ_Recurrente.Range("B2").Value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_EJ_Recurrente.Range("C999").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Integer
    For i = 9 To r
        wshGL_EJ_Recurrente.Range("C" & rowToBeUsed).Value = JENo
        wshGL_EJ_Recurrente.Range("D" & rowToBeUsed).Value = wshGL_EJ.Range("F6").Value
        wshGL_EJ_Recurrente.Range("E" & rowToBeUsed).Value = wshGL_EJ.Range("L" & i).Value
        wshGL_EJ_Recurrente.Range("F" & rowToBeUsed).Value = wshGL_EJ.Range("E" & i).Value
        If wshGL_EJ.Range("H" & i).Value <> "" Then
            wshGL_EJ_Recurrente.Range("G" & rowToBeUsed).Value = wshGL_EJ.Range("H" & i).Value
        End If
        If wshGL_EJ.Range("I" & i).Value <> "" Then
            wshGL_EJ_Recurrente.Range("H" & rowToBeUsed).Value = wshGL_EJ.Range("I" & i).Value
        End If
        wshGL_EJ_Recurrente.Range("I" & rowToBeUsed).Value = wshGL_EJ.Range("J" & i).Value
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call GL_EJ_Auto_Build_Summary '2024-03-14 @ 07:40
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("modGL_JE:GL_EJ_Auto_Add_Record_Locally()", timerStart)
    
End Sub

