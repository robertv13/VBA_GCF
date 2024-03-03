Attribute VB_Name = "modJE"
Option Explicit

Sub JE_Update()

    If IsDateValide = False Then Exit Sub
    
    If IsEcritureBalance = False Then Exit Sub
    
    Dim rowEJLast As Long
    rowEJLast = wshJE.Range("D23").End(xlUp).row  'Last Used Row in wshJE
    If IsEcritureValide(rowEJLast) = False Then Exit Sub
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    Call Add_GL_Trans_Record_To_DB(rowEJLast)
    Call Add_GL_Trans_Record_Locally(rowEJLast)
    
    If wshJE.ckbRecurrente = True Then
        Save_EJ_Recurrente rowEJLast
    End If
    
    With wshJE
        'Increment Next JE number
        .Range("B1").value = .Range("B1").value + 1
        
        Call wshJE_Clear_All_Cells
        
        wshJE.Activate
        .Range("E4").Select
        .Range("E4").Activate
    End With
    
    MsgBox "L'écriture a été reporté avec succès"
    
End Sub

Sub Save_EJ_Recurrente(ll As Long)

    Dim rowEJLast As Long
    rowEJLast = wshJE.Range("D99").End(xlUp).row  'Last Used Row in wshJE
    
    Call Add_JE_Auto_Record_To_DB(ll)
    Call Add_JE_Auto_Record_Locally(ll)
    
End Sub

Sub LoadJEAutoIntoJE(EJAutoDesc As String, NoEJAuto As Long)

    'On copie l'E/J automatique vers wshEJ
    Dim rowJEAuto, rowJE As Long
    rowJEAuto = wshEJRecurrente.Range("C99999").End(xlUp).row  'Last Row used in wshJERecuurente
    
    Call wshJE_Clear_All_Cells
    rowJE = 9
    
    Dim r As Long
    For r = 2 To rowJEAuto
        If wshEJRecurrente.Range("C" & r).value = NoEJAuto And wshEJRecurrente.Range("E" & r).value <> "" Then
            wshJE.Range("D" & rowJE).value = wshEJRecurrente.Range("F" & r).value
            wshJE.Range("G" & rowJE).value = wshEJRecurrente.Range("G" & r).value
            wshJE.Range("H" & rowJE).value = wshEJRecurrente.Range("H" & r).value
            wshJE.Range("I" & rowJE).value = wshEJRecurrente.Range("I" & r).value
            wshJE.Range("K" & rowJE).value = wshEJRecurrente.Range("E" & r).value
            rowJE = rowJE + 1
        End If
    Next r
    wshJE.Range("E6").value = "Auto - " & EJAutoDesc
    wshJE.Range("J4").Activate

End Sub

Sub wshJE_Clear_All_Cells()

    'Efface toutes les cellules de la feuille
    Application.EnableEvents = False
    With wshJE
        .Range("E4,J4,E6:J6").ClearContents
        .Range("D9:F23,G9:G23,H9:H23,I9:J23,K9:K23").ClearContents
        .ckbRecurrente = False
        .Range("J4").value = " "
    Application.EnableEvents = True
    End With

End Sub

Function IsDateValide() As Boolean

    IsDateValide = False
    If wshJE.Range("J4").value = "" Or IsDate(wshJE.Range("J4").value) = False Then
        MsgBox "Une date d'écriture est obligatoire." & vbNewLine & vbNewLine & _
            "Veuillez saisir une date valide!", vbCritical, "Date Invalide"
    Else
        IsDateValide = True
    End If

End Function

Function IsEcritureBalance() As Boolean

    IsEcritureBalance = False
    If wshJE.Range("G25").value <> wshJE.Range("H25").value Then
        MsgBox "Votre écriture ne balance pas." & vbNewLine & vbNewLine & _
            "Débits = " & wshJE.Range("G25").value & " et Crédits = " & wshJE.Range("H25").value & vbNewLine & vbNewLine & _
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
        If wshJE.Range("D" & i).value <> "" Then
            If wshJE.Range("G" & i).value = "" And wshJE.Range("H" & i).value = "" Then
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
    strSQL = "SELECT MAX(No_Entrée) AS MaxEJNo FROM [" & sheetName & "$]"

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
            rs.Fields("No_Entrée").value = nextJENo
            rs.Fields("Date").value = CDate(wshJE.Range("J4").value)
            rs.Fields("Description").value = wshJE.Range("E6").value
            rs.Fields("Source").value = wshJE.Range("E4").value
            rs.Fields("No_Compte").value = wshJE.Range("K" & l).value
            rs.Fields("Compte").value = wshJE.Range("D" & l).value
            rs.Fields("Débit").value = wshJE.Range("G" & l).value
            rs.Fields("Crédit").value = wshJE.Range("H" & l).value
            rs.Fields("AutreRemarque").value = wshJE.Range("I" & l).value
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
        wshGL_Trans.Range("B" & rowToBeUsed).value = CDate(wshJE.Range("J4").value)
        wshGL_Trans.Range("C" & rowToBeUsed).value = wshJE.Range("E6").value
        wshGL_Trans.Range("D" & rowToBeUsed).value = wshJE.Range("E4").value
        wshGL_Trans.Range("E" & rowToBeUsed).value = wshJE.Range("K" & i).value
        wshGL_Trans.Range("F" & rowToBeUsed).value = wshJE.Range("D" & i).value
        If wshJE.Range("G" & i).value <> "" Then
            wshGL_Trans.Range("G" & rowToBeUsed).value = wshJE.Range("G" & i).value
        End If
        If wshJE.Range("H" & i).value <> "" Then
            wshGL_Trans.Range("H" & rowToBeUsed).value = wshJE.Range("H" & i).value
        End If
        wshGL_Trans.Range("I" & rowToBeUsed).value = wshJE.Range("I" & i).value
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
            rs.Fields("Description").value = wshJE.Range("E6").value
            rs.Fields("No_Compte").value = wshJE.Range("K" & l).value
            rs.Fields("Compte").value = wshJE.Range("D" & l).value
            rs.Fields("Débit").value = wshJE.Range("G" & l).value
            rs.Fields("Crédit").value = wshJE.Range("H" & l).value
            rs.Fields("AutreRemarque").value = wshJE.Range("I" & l).value
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
    JENo = wshEJRecurrente.Range("B1").value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshEJRecurrente.Range("C999").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Integer
    For i = 9 To r
        wshEJRecurrente.Range("C" & rowToBeUsed).value = JENo
        wshEJRecurrente.Range("D" & rowToBeUsed).value = wshJE.Range("E6").value
        wshEJRecurrente.Range("E" & rowToBeUsed).value = wshJE.Range("K" & i).value
        wshEJRecurrente.Range("F" & rowToBeUsed).value = wshJE.Range("D" & i).value
        If wshJE.Range("G" & i).value <> "" Then
            wshEJRecurrente.Range("G" & rowToBeUsed).value = wshJE.Range("G" & i).value
        End If
        If wshJE.Range("H" & i).value <> "" Then
            wshEJRecurrente.Range("H" & rowToBeUsed).value = wshJE.Range("H" & i).value
        End If
        wshEJRecurrente.Range("I" & rowToBeUsed).value = wshJE.Range("I" & i).value
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("Add_JE_Auto_Record_Locally()", timerStart)
    
End Sub

Sub UpdateJEAuto()
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call GLJEAuto_Import
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
  
End Sub

Sub GLJEAuto_Import() '2024-01-07 @ 14:45
    
    Application.ScreenUpdating = False
    
    Dim saveLastRow As Long
    saveLastRow = wshEJRecurrente.Range("C999999").End(xlUp).row + 1
    
    'Clear all cells, but the headers, in the target worksheet
    wshEJRecurrente.Range("C1").CurrentRegion.Offset(1, 0).ClearContents
    wshEJRecurrente.Range("L1").CurrentRegion.Offset(1, 0).ClearContents

    'Import JEAuto from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-01-07
                     
    'Set up source and destination ranges
    Dim sourceRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets("EJ_Auto").usedRange

    Dim destinationRange As Range
    Set destinationRange = wshEJRecurrente.Range("C1")

    'Copy data, using Range to Range
    sourceRange.Copy destinationRange
    wshEJRecurrente.Range("C1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False
    
    'Fill the list of Automatic J/E (same worksheet)
    With wshEJRecurrente
        Dim i As Long, rsomm As Long, oldNOEJA As String
        rsomm = 2
        For i = 2 To .Range("C9999").End(xlUp).row
            If .Range("C" & i).value <> oldNOEJA Then
                .Range("L" & rsomm).value = .Range("D" & i).value
                .Range("M" & rsomm).value = .Range("C" & i).value
                oldNOEJA = .Range("C" & i).value
                rsomm = rsomm + 1
            End If
        Next i
    End With
    
    Application.ScreenUpdating = True
    
End Sub


