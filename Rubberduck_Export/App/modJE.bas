Attribute VB_Name = "modJE"
Option Explicit

Sub JE_Update()

    If IsDateValide = False Then Exit Sub
    
    If IsEcritureBalance = False Then Exit Sub
    
    Dim rowEJLast As Long
    rowEJLast = wshJE.Range("E23").End(xlUp).row  'Last Used Row in wshJE
    If IsEcritureValide(rowEJLast) = False Then Exit Sub
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    Call Add_GL_Trans_Record_To_DB(rowEJLast)
    Call Add_GL_Trans_Record_Locally(rowEJLast)
    
    If wshJE.ckbRecurrente = True Then
        Call Save_EJ_Recurrente(rowEJLast)
    End If
    
    'Save Current JE number
    Dim strCurrentJE As String
    strCurrentJE = wshJE.Range("B1").Value
    
    'Increment Next JE number
    wshJE.Range("B1").Value = wshJE.Range("B1").Value + 1
        
    Call wshJE_Clear_All_Cells
        
    With wshJE
        .Activate
        .Range("F4").Select
        .Range("F4").Activate
    End With
    
    MsgBox "L'écriture numéro '" & strCurrentJE & "' a été reporté avec succès"
    
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
        If wshEJRecurrente.Range("C" & r).Value = NoEJAuto And wshEJRecurrente.Range("E" & r).Value <> "" Then
            wshJE.Range("E" & rowJE).Value = wshEJRecurrente.Range("F" & r).Value
            wshJE.Range("H" & rowJE).Value = wshEJRecurrente.Range("G" & r).Value
            wshJE.Range("I" & rowJE).Value = wshEJRecurrente.Range("H" & r).Value
            wshJE.Range("J" & rowJE).Value = wshEJRecurrente.Range("I" & r).Value
            wshJE.Range("L" & rowJE).Value = wshEJRecurrente.Range("E" & r).Value
            rowJE = rowJE + 1
        End If
    Next r
    wshJE.Range("F6").Value = "[Auto]-" & EJAutoDesc
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
    If wshJE.Range("K4").Value = "" Or IsDate(wshJE.Range("K4").Value) = False Then
        MsgBox "Une date d'écriture est obligatoire." & vbNewLine & vbNewLine & _
            "Veuillez saisir une date valide!", vbCritical, "Date Invalide"
    Else
        IsDateValide = True
    End If

End Function

Function IsEcritureBalance() As Boolean

    IsEcritureBalance = False
    If wshJE.Range("H26").Value <> wshJE.Range("I26").Value Then
        MsgBox "Votre écriture ne balance pas." & vbNewLine & vbNewLine & _
            "Débits = " & wshJE.Range("H26").Value & " et Crédits = " & wshJE.Range("I26").Value & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée.", vbCritical, "Veuillez vérifier votre écriture!"
    Else
        IsEcritureBalance = True
    End If

End Function

Function IsEcritureValide(rmax As Long) As Boolean

    IsEcritureValide = True 'Optimist
    If rmax <= 9 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        IsEcritureValide = False
    End If
    
    Dim i As Long
    For i = 9 To rmax
        If wshJE.Range("E" & i).Value <> "" Then
            If wshJE.Range("H" & i).Value = "" And wshJE.Range("I" & i).Value = "" Then
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
    fullFileName = wshAdmin.Range("FolderSharedData").Value & _
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
    strSQL = "SELECT MAX(No_Entrée) AS MaxEJNo FROM [" & sheetName & "$]"

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
    wshJE.Range("B1").Value = nextJENo
    
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
            rs.Fields("No_Entrée").Value = nextJENo
            rs.Fields("Date").Value = CDate(wshJE.Range("K4").Value)
            rs.Fields("Description").Value = wshJE.Range("F6").Value
            rs.Fields("Source").Value = wshJE.Range("F4").Value
            rs.Fields("No_Compte").Value = wshJE.Range("L" & l).Value
            rs.Fields("Compte").Value = wshJE.Range("E" & l).Value
            rs.Fields("Débit").Value = wshJE.Range("H" & l).Value
            rs.Fields("Crédit").Value = wshJE.Range("I" & l).Value
            rs.Fields("AutreRemarque").Value = wshJE.Range("J" & l).Value
            rs.Fields("TimeStamp").Value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
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
    JENo = wshJE.Range("B1").Value
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_Trans.Range("A99999").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Integer
    For i = 9 To r
        wshGL_Trans.Range("A" & rowToBeUsed).Value = JENo
        wshGL_Trans.Range("B" & rowToBeUsed).Value = CDate(wshJE.Range("K4").Value)
        wshGL_Trans.Range("C" & rowToBeUsed).Value = wshJE.Range("F6").Value
        wshGL_Trans.Range("D" & rowToBeUsed).Value = wshJE.Range("F4").Value
        wshGL_Trans.Range("E" & rowToBeUsed).Value = wshJE.Range("L" & i).Value
        wshGL_Trans.Range("F" & rowToBeUsed).Value = wshJE.Range("E" & i).Value
        If wshJE.Range("H" & i).Value <> "" Then
            wshGL_Trans.Range("G" & rowToBeUsed).Value = wshJE.Range("H" & i).Value
        End If
        If wshJE.Range("I" & i).Value <> "" Then
            wshGL_Trans.Range("H" & rowToBeUsed).Value = wshJE.Range("I" & i).Value
        End If
        wshGL_Trans.Range("I" & rowToBeUsed).Value = wshJE.Range("J" & i).Value
        wshGL_Trans.Range("J" & rowToBeUsed).Value = CDate(Now())
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call Output_Timer_Results("Add_GL_Trans_Record_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub Add_JE_Auto_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").Value & _
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
    If IsNull(rs.Fields("MaxEJANo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastEJA = 1
    Else
        lastEJA = rs.Fields("MaxEJANo").Value
    End If
    
    'Calculate the new ID
    nextEJANo = lastEJA + 1
    wshEJRecurrente.Range("B2").Value = nextEJANo

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("No_EJA").Value = nextEJANo
            rs.Fields("Description").Value = wshJE.Range("F6").Value
            rs.Fields("No_Compte").Value = wshJE.Range("L" & l).Value
            rs.Fields("Compte").Value = wshJE.Range("E" & l).Value
            rs.Fields("Débit").Value = wshJE.Range("H" & l).Value
            rs.Fields("Crédit").Value = wshJE.Range("I" & l).Value
            rs.Fields("AutreRemarque").Value = wshJE.Range("J" & l).Value
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
    JENo = wshEJRecurrente.Range("B2").Value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshEJRecurrente.Range("C999").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Integer
    For i = 9 To r
        wshEJRecurrente.Range("C" & rowToBeUsed).Value = JENo
        wshEJRecurrente.Range("D" & rowToBeUsed).Value = wshJE.Range("F6").Value
        wshEJRecurrente.Range("E" & rowToBeUsed).Value = wshJE.Range("L" & i).Value
        wshEJRecurrente.Range("F" & rowToBeUsed).Value = wshJE.Range("E" & i).Value
        If wshJE.Range("H" & i).Value <> "" Then
            wshEJRecurrente.Range("G" & rowToBeUsed).Value = wshJE.Range("H" & i).Value
        End If
        If wshJE.Range("I" & i).Value <> "" Then
            wshEJRecurrente.Range("H" & rowToBeUsed).Value = wshJE.Range("I" & i).Value
        End If
        wshEJRecurrente.Range("I" & rowToBeUsed).Value = wshJE.Range("J" & i).Value
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("Add_JE_Auto_Record_Locally()", timerStart)
    
End Sub

