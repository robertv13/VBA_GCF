Attribute VB_Name = "modGL_EJ"
Option Explicit

Sub JE_Update()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:JE_Update", 0)
    
    If Fn_Is_Date_Valide(wshGL_EJ.Range("K4").value) = False Then Exit Sub
    
    If Fn_Is_Ecriture_Balance = False Then Exit Sub
    
    Dim rowEJLast As Long
    rowEJLast = wshGL_EJ.Range("E23").End(xlUp).Row  'Last Used Row in wshGL_EJ
    If Fn_Is_JE_Valid(rowEJLast) = False Then Exit Sub
    
    'Transfert des donn�es vers wshGL, ent�te d'abord puis une ligne � la fois
    Call GL_Trans_Add_Record_To_DB(rowEJLast)
    Call GL_Trans_Add_Record_Locally(rowEJLast)
    
    If wshGL_EJ.ckbRecurrente = True Then
        Call Save_EJ_Recurrente(rowEJLast)
    End If
    
    'Save Current JE number
    Dim strCurrentJE As String
    strCurrentJE = wshGL_EJ.Range("B1").value
    
    'Increment Next JE number
    wshGL_EJ.Range("B1").value = wshGL_EJ.Range("B1").value + 1
        
    Call wshGL_EJ_Clear_All_Cells
        
    With wshGL_EJ
        .Activate
        .Range("F4").Select
        .Range("F4").Activate
    End With
    
    MsgBox "L'�criture num�ro '" & strCurrentJE & "' a �t� report� avec succ�s"
    
    Call Log_Record("modGL_EJ:JE_Update", startTime)
    
End Sub

Sub Save_EJ_Recurrente(ll As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:Save_EJ_Recurrente", 0)
    
    Dim rowEJLast As Long
    rowEJLast = wshGL_EJ.Range("E99").End(xlUp).Row  'Last Used Row in wshGL_EJ
    
    Call GL_EJ_Auto_Add_Record_To_DB(ll)
    Call GL_EJ_Auto_Add_Record_Locally(ll)
    
    Call Log_Record("modGL_EJ:Save_EJ_Recurrente", startTime)
    
End Sub

Sub Load_JEAuto_Into_JE(EJAutoDesc As String, NoEJAuto As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:Load_JEAuto_Into_JE", 0)
    
    'On copie l'E/J automatique vers wshEJ
    Dim rowJEAuto, rowJE As Long
    rowJEAuto = wshGL_EJ_Recurrente.Cells(wshGL_EJ_Recurrente.rows.count, "A").End(xlUp).Row  'Last Row used in wshGL_EJRecuurente
    
    Call wshGL_EJ_Clear_All_Cells
    rowJE = 9
    
    Dim r As Long
    For r = 2 To rowJEAuto
        If wshGL_EJ_Recurrente.Range("A" & r).value = NoEJAuto And wshGL_EJ_Recurrente.Range("C" & r).value <> "" Then
            wshGL_EJ.Range("E" & rowJE).value = wshGL_EJ_Recurrente.Range("D" & r).value
            wshGL_EJ.Range("H" & rowJE).value = wshGL_EJ_Recurrente.Range("E" & r).value
            wshGL_EJ.Range("I" & rowJE).value = wshGL_EJ_Recurrente.Range("F" & r).value
            wshGL_EJ.Range("J" & rowJE).value = wshGL_EJ_Recurrente.Range("G" & r).value
            wshGL_EJ.Range("L" & rowJE).value = wshGL_EJ_Recurrente.Range("C" & r).value
            rowJE = rowJE + 1
        End If
    Next r
    wshGL_EJ.Range("F6").value = "[Auto]-" & EJAutoDesc
    wshGL_EJ.Range("K4").Activate

    Call Log_Record("modGL_EJ:Load_JEAuto_Into_JE", startTime)
    
End Sub

Sub wshGL_EJ_Clear_All_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:wshGL_EJ_Clear_All_Cells", 0)
    
    'Efface toutes les cellules de la feuille
    Application.EnableEvents = False
    ActiveSheet.Unprotect
    With wshGL_EJ
        .Range("F4,F6:K6").ClearContents
        .Range("E9:G23,H9:H23,I9:I23,J9:L23").ClearContents
        .ckbRecurrente = False
        Application.EnableEvents = True
        wshGL_EJ.Activate
        wshGL_EJ.Range("F4").Select
    End With
    
    With ActiveSheet
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Call Log_Record("modGL_EJ:wshGL_EJ_Clear_All_Cells", startTime)

End Sub

Sub GL_EJ_Auto_Build_Summary()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Auto_Build_Summary", 0)
    
    'Build the summary at column K & L
    Dim lastUsedRow1 As Long
    lastUsedRow1 = wshGL_EJ_Recurrente.Cells(wshGL_EJ_Recurrente.rows.count, "A").End(xlUp).Row
    
    Dim lastUsedRow2 As Long
    lastUsedRow2 = wshGL_EJ_Recurrente.Cells(wshGL_EJ_Recurrente.rows.count, "I").End(xlUp).Row
    If lastUsedRow2 > 1 Then
        wshGL_EJ_Recurrente.Range("I2:J" & lastUsedRow2).ClearContents
    End If
    
    With wshGL_EJ_Recurrente
        Dim i As Long, k As Long, oldEntry As String
        k = 2
        For i = 2 To lastUsedRow1
            If .Range("A" & i).value <> oldEntry Then
                .Range("I" & k).value = .Range("B" & i).value
                .Range("J" & k).value = "'" & Fn_Pad_A_String(.Range("A" & i).value, " ", 5, "L")
                oldEntry = .Range("A" & i).value
                k = k + 1
            End If
        Next i
    End With

    Call Log_Record("modGL_EJ:GL_EJ_Auto_Build_Summary", startTime)

End Sub

Sub GL_Trans_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_Trans_Add_Record_To_DB", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(No_Entr�e) AS MaxEJNo FROM [" & destinationTab & "$]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim MaxEJNo As Long, lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new JE number
    Dim nextJENo As Long
    nextJENo = lastJE + 1
    wshGL_EJ.Range("B1").value = nextJENo
    
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
            rs.Fields("No_Entr�e").value = nextJENo
            rs.Fields("Date").value = CDate(wshGL_EJ.Range("K4").value)
            rs.Fields("Description").value = wshGL_EJ.Range("F6").value
            rs.Fields("Source").value = wshGL_EJ.Range("F4").value
            rs.Fields("No_Compte").value = wshGL_EJ.Range("L" & l).value
            rs.Fields("Compte").value = wshGL_EJ.Range("E" & l).value
            rs.Fields("D�bit").value = wshGL_EJ.Range("H" & l).value
            rs.Fields("Cr�dit").value = wshGL_EJ.Range("I" & l).value
            rs.Fields("AutreRemarque").value = wshGL_EJ.Range("J" & l).value
            rs.Fields("TimeStamp").value = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
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
    
    Call Log_Record("modGL_EJ:GL_Trans_Add_Record_To_DB", startTime)

End Sub

Sub GL_Trans_Add_Record_Locally(r As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshGL_EJ.Range("B1").value
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_Trans.Range("A99999").End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Long
    For i = 9 To r
        wshGL_Trans.Range("A" & rowToBeUsed).value = JENo
        wshGL_Trans.Range("B" & rowToBeUsed).value = CDate(wshGL_EJ.Range("K4").value)
        wshGL_Trans.Range("C" & rowToBeUsed).value = wshGL_EJ.Range("F6").value
        wshGL_Trans.Range("D" & rowToBeUsed).value = wshGL_EJ.Range("F4").value
        wshGL_Trans.Range("E" & rowToBeUsed).value = wshGL_EJ.Range("L" & i).value
        wshGL_Trans.Range("F" & rowToBeUsed).value = wshGL_EJ.Range("E" & i).value
        If wshGL_EJ.Range("H" & i).value <> "" Then
            wshGL_Trans.Range("G" & rowToBeUsed).value = wshGL_EJ.Range("H" & i).value
        End If
        If wshGL_EJ.Range("I" & i).value <> "" Then
            wshGL_Trans.Range("H" & rowToBeUsed).value = wshGL_EJ.Range("I" & i).value
        End If
        wshGL_Trans.Range("I" & rowToBeUsed).value = wshGL_EJ.Range("J" & i).value
        wshGL_Trans.Range("J" & rowToBeUsed).value = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally", startTime)

End Sub

Sub GL_EJ_Auto_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Auto_Add_Record_To_DB", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_EJ_Auto"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJANo As Long
    strSQL = "SELECT MAX(No_EJA) AS MaxEJANo FROM [" & destinationTab & "$]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastEJA As Long, nextEJANo As Long
    If IsNull(rs.Fields("MaxEJANo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastEJA = 1
    Else
        lastEJA = rs.Fields("MaxEJANo").value
    End If
    
    'Calculate the new ID
    nextEJANo = lastEJA + 1
    wshGL_EJ_Recurrente.Range("B2").value = nextEJANo

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("No_EJA").value = nextEJANo
            rs.Fields("Description").value = wshGL_EJ.Range("F6").value
            rs.Fields("No_Compte").value = wshGL_EJ.Range("L" & l).value
            rs.Fields("Compte").value = wshGL_EJ.Range("E" & l).value
            rs.Fields("D�bit").value = wshGL_EJ.Range("H" & l).value
            rs.Fields("Cr�dit").value = wshGL_EJ.Range("I" & l).value
            rs.Fields("AutreRemarque").value = wshGL_EJ.Range("J" & l).value
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
    
    Call Log_Record("modGL_EJ:GL_EJ_Auto_Add_Record_To_DB", startTime)

End Sub

Sub GL_EJ_Auto_Add_Record_Locally(r As Long) 'Write records to local file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Auto_Add_Record_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshGL_EJ_Recurrente.Range("B2").value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_EJ_Recurrente.Range("C999").End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Long
    For i = 9 To r
        wshGL_EJ_Recurrente.Range("C" & rowToBeUsed).value = JENo
        wshGL_EJ_Recurrente.Range("D" & rowToBeUsed).value = wshGL_EJ.Range("F6").value
        wshGL_EJ_Recurrente.Range("E" & rowToBeUsed).value = wshGL_EJ.Range("L" & i).value
        wshGL_EJ_Recurrente.Range("F" & rowToBeUsed).value = wshGL_EJ.Range("E" & i).value
        If wshGL_EJ.Range("H" & i).value <> "" Then
            wshGL_EJ_Recurrente.Range("G" & rowToBeUsed).value = wshGL_EJ.Range("H" & i).value
        End If
        If wshGL_EJ.Range("I" & i).value <> "" Then
            wshGL_EJ_Recurrente.Range("H" & rowToBeUsed).value = wshGL_EJ.Range("I" & i).value
        End If
        wshGL_EJ_Recurrente.Range("I" & rowToBeUsed).value = wshGL_EJ.Range("J" & i).value
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call GL_EJ_Auto_Build_Summary '2024-03-14 @ 07:40
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modGL_EJ:GL_EJ_Auto_Add_Record_Locally", startTime)
    
End Sub


