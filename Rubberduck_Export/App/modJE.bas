Attribute VB_Name = "modJE"
Option Explicit

Sub JE_Update()

    If IsDateValide = False Then Exit Sub
    
    If IsEcritureBalance = False Then Exit Sub
    
    Dim rowEJLast As Long
    rowEJLast = wshJE.Range("E23").End(xlUp).row  'Last Used Row in wshJE
    If IsEcritureValide(rowEJLast) = False Then Exit Sub
    
    'Transfert des donn�es vers wshGL, ent�te d'abord puis une ligne � la fois
    Call Add_GL_Trans_Record_To_DB(rowEJLast)
    Call Add_GL_Trans_Record_Locally(rowEJLast)
    
    If wshJE.ckbRecurrente = True Then
        Call Save_EJ_Recurrente(rowEJLast)
    End If
    
    'Save Current JE number
    Dim strCurrentJE As String
    strCurrentJE = wshJE.Range("B1").value
    
    'Increment Next JE number
    wshJE.Range("B1").value = wshJE.Range("B1").value + 1
        
    Call wshJE_Clear_All_Cells
        
    With wshJE
        .Activate
        .Range("F4").Select
        .Range("F4").Activate
    End With
    
    MsgBox "L'�criture num�ro '" & strCurrentJE & "' a �t� report� avec succ�s"
    
End Sub

Sub Save_EJ_Recurrente(ll As Long)

    Dim rowEJLast As Long
    rowEJLast = wshJE.Range("E99").End(xlUp).row  'Last Used Row in wshJE
    
    Call Add_JE_Auto_Record_To_DB(ll)
    Call Add_JE_Auto_Record_Locally(ll)
    
End Sub

Sub Load_JEAuto_Into_JE(EJAutoDesc As String, NoEJAuto As Long)

    'On copie l'E/J automatique vers wshEJ
    Dim rowJEAuto, rowJE As Long
    rowJEAuto = wshEJRecurrente.Range("C99999").End(xlUp).row  'Last Row used in wshJERecuurente
    
    Call wshJE_Clear_All_Cells
    rowJE = 9
    
    Dim r As Long
    For r = 2 To rowJEAuto
        If wshEJRecurrente.Range("C" & r).value = NoEJAuto And wshEJRecurrente.Range("E" & r).value <> "" Then
            wshJE.Range("E" & rowJE).value = wshEJRecurrente.Range("F" & r).value
            wshJE.Range("H" & rowJE).value = wshEJRecurrente.Range("G" & r).value
            wshJE.Range("I" & rowJE).value = wshEJRecurrente.Range("H" & r).value
            wshJE.Range("J" & rowJE).value = wshEJRecurrente.Range("I" & r).value
            wshJE.Range("L" & rowJE).value = wshEJRecurrente.Range("E" & r).value
            rowJE = rowJE + 1
        End If
    Next r
    wshJE.Range("F6").value = "[Auto]-" & EJAutoDesc
    wshJE.Range("K4").Activate

End Sub

Sub wshJE_Clear_All_Cells()

    'Efface toutes les cellules de la feuille
    Application.EnableEvents = False
    With wshJE
        .Range("F4,F6:K6").ClearContents
        .Range("E9:G23,H9:H23,I9:I23,J9:L23").ClearContents
        .ckbRecurrente = False
    Application.EnableEvents = True
    wshJE.Range("F4").Activate
    End With

End Sub

Function IsDateValide() As Boolean

    IsDateValide = False
    If wshJE.Range("K4").value = "" Or IsDate(wshJE.Range("K4").value) = False Then
        MsgBox "Une date d'�criture est obligatoire." & vbNewLine & vbNewLine & _
            "Veuillez saisir une date valide!", vbCritical, "Date Invalide"
    Else
        IsDateValide = True
    End If

End Function

Function IsEcritureBalance() As Boolean

    IsEcritureBalance = False
    If wshJE.Range("H26").value <> wshJE.Range("I26").value Then
        MsgBox "Votre �criture ne balance pas." & vbNewLine & vbNewLine & _
            "D�bits = " & wshJE.Range("H26").value & " et Cr�dits = " & wshJE.Range("I26").value & vbNewLine & vbNewLine & _
            "Elle n'est donc pas report�e.", vbCritical, "Veuillez v�rifier votre �criture!"
    Else
        IsEcritureBalance = True
    End If

End Function

Function IsEcritureValide(rmax As Long) As Boolean

    IsEcritureValide = True 'Optimist
    If rmax <= 9 Or rmax > 23 Then
        MsgBox "L'�criture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas report�e!", vbCritical, "Vous devez v�rifier l'�criture"
        IsEcritureValide = False
    End If
    
    Dim i As Long
    For i = 9 To rmax
        If wshJE.Range("E" & i).value <> "" Then
            If wshJE.Range("H" & i).value = "" And wshJE.Range("I" & i).value = "" Then
                MsgBox "Il existe une ligne avec un compte, sans montant !"
                IsEcritureValide = False
            End If
        End If
    Next i

End Function

Sub Add_GL_Trans_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & _
                   Application.PathSeparator & "GCF_BD_Sortie.xlsx"
    sheetName = "GL_Trans"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(No_Entr�e) AS MaxEJNo FROM [" & sheetName & "$]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim maxEJNo As Long, lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new JE number
    Dim nextJENo As Long
    nextJENo = lastJE + 1
    wshJE.Range("B1").value = nextJENo
    
    'Build formula
    Dim formula As String
    formula = "=ROW()"

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
    
    'Read all line from Journal Entry
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("No_Entr�e").value = nextJENo
            rs.Fields("Date").value = CDate(wshJE.Range("K4").value)
            rs.Fields("Description").value = wshJE.Range("F6").value
            rs.Fields("Source").value = wshJE.Range("F4").value
            rs.Fields("No_Compte").value = wshJE.Range("L" & l).value
            rs.Fields("Compte").value = wshJE.Range("E" & l).value
            rs.Fields("D�bit").value = wshJE.Range("H" & l).value
            rs.Fields("Cr�dit").value = wshJE.Range("I" & l).value
            rs.Fields("AutreRemarque").value = wshJE.Range("J" & l).value
            rs.Fields("TimeStamp").value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        rs.update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("Add_GL_Trans_Record_To_DB()", timerStart)

End Sub

Sub Add_GL_Trans_Record_Locally(r As Long) 'Write records locally
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshJE.Range("B1").value
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_Trans.Range("A99999").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Integer
    For i = 9 To r
        wshGL_Trans.Range("A" & rowToBeUsed).value = JENo
        wshGL_Trans.Range("B" & rowToBeUsed).value = CDate(wshJE.Range("K4").value)
        wshGL_Trans.Range("C" & rowToBeUsed).value = wshJE.Range("F6").value
        wshGL_Trans.Range("D" & rowToBeUsed).value = wshJE.Range("F4").value
        wshGL_Trans.Range("E" & rowToBeUsed).value = wshJE.Range("L" & i).value
        wshGL_Trans.Range("F" & rowToBeUsed).value = wshJE.Range("E" & i).value
        If wshJE.Range("H" & i).value <> "" Then
            wshGL_Trans.Range("G" & rowToBeUsed).value = wshJE.Range("H" & i).value
        End If
        If wshJE.Range("I" & i).value <> "" Then
            wshGL_Trans.Range("H" & rowToBeUsed).value = wshJE.Range("I" & i).value
        End If
        wshGL_Trans.Range("I" & rowToBeUsed).value = wshJE.Range("J" & i).value
        wshGL_Trans.Range("J" & rowToBeUsed).value = CDate(Now())
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call Output_Timer_Results("Add_GL_Trans_Record_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub Add_JE_Auto_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & _
                   Application.PathSeparator & "GCF_BD_Sortie.xlsx"
    sheetName = "EJ_Auto"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJANo As Long
    strSQL = "SELECT MAX(No_EJA) AS MaxEJANo FROM [" & sheetName & "$]"

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
    wshEJRecurrente.Range("B2").value = nextEJANo

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("No_EJA").value = nextEJANo
            rs.Fields("Description").value = wshJE.Range("F6").value
            rs.Fields("No_Compte").value = wshJE.Range("L" & l).value
            rs.Fields("Compte").value = wshJE.Range("E" & l).value
            rs.Fields("D�bit").value = wshJE.Range("H" & l).value
            rs.Fields("Cr�dit").value = wshJE.Range("I" & l).value
            rs.Fields("AutreRemarque").value = wshJE.Range("J" & l).value
        rs.update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("Add_JE_Auto_Record_To_DB()", timerStart)

End Sub

Sub Add_JE_Auto_Record_Locally(r As Long) 'Write records to local file
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshEJRecurrente.Range("B2").value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshEJRecurrente.Range("C999").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Integer
    For i = 9 To r
        wshEJRecurrente.Range("C" & rowToBeUsed).value = JENo
        wshEJRecurrente.Range("D" & rowToBeUsed).value = wshJE.Range("F6").value
        wshEJRecurrente.Range("E" & rowToBeUsed).value = wshJE.Range("L" & i).value
        wshEJRecurrente.Range("F" & rowToBeUsed).value = wshJE.Range("E" & i).value
        If wshJE.Range("H" & i).value <> "" Then
            wshEJRecurrente.Range("G" & rowToBeUsed).value = wshJE.Range("H" & i).value
        End If
        If wshJE.Range("I" & i).value <> "" Then
            wshEJRecurrente.Range("H" & rowToBeUsed).value = wshJE.Range("I" & i).value
        End If
        wshEJRecurrente.Range("I" & rowToBeUsed).value = wshJE.Range("J" & i).value
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("Add_JE_Auto_Record_Locally()", timerStart)
    
End Sub

