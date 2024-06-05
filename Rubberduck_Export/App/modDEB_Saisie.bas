Attribute VB_Name = "modDEB_Saisie"
Option Explicit

Sub DEB_Saisie_Update()

    If Fn_Is_Date_Valide(wshDEB_Saisie.Range("O4").value) = False Then Exit Sub
    
    If Fn_Is_Debours_Balance = False Then Exit Sub
    
    Dim rowDebSaisie As Long
    rowDebSaisie = wshDEB_Saisie.Range("E23").End(xlUp).row  'Last Used Row in wshDEB_Saisie
    If Fn_Is_Deb_Saisie_Valid(rowDebSaisie) = False Then Exit Sub
    
    'Transfert des données vers Débours_Trans, entête d'abord puis une ligne à la fois
    Call DEB_Trans_Add_Record_To_DB(rowDebSaisie)
    Call DEB_Trans_Add_Record_Locally(rowDebSaisie)
    
'    If wshDEB_Saisie.ckbRecurrente = True Then
'        Call Save_EJ_Recurrente(rowEJLast)
'    End If
    
    'GL posting
    Call DEB_Saisie_GL_Posting_Preparation
    
    'Save Current DEBOURS number
    Dim CurrentDeboursNo As String
    CurrentDeboursNo = wshDEB_Saisie.Range("B1").value
    
    MsgBox "Le déboursé, numéro '" & CurrentDeboursNo & "' a été reporté avec succès"
    
    Call DEB_Saisie_Clear_All_Cells
        
End Sub

Sub DEB_Trans_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "DEB_Trans"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(No_Entrée) AS MaxDebTransNo FROM [" & destinationTab & "$]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim maxDebTransNo As Long, lastDebTrans As Long
    If IsNull(rs.Fields("MaxDebTransNo").value) Then
        'Handle empty table (assign a default value, e.g., 0)
        lastDebTrans = 0
    Else
        lastDebTrans = rs.Fields("MaxDebTransNo").value
    End If
    
    'Calculate the new JE number
    Dim nextDebTransNo As Long
    nextDebTransNo = lastDebTrans + 1
    Application.EnableEvents = False
    wshDEB_Saisie.Range("B1").value = nextDebTransNo
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
            rs.Fields("No_Entrée").value = nextDebTransNo
            rs.Fields("Date").value = CDate(wshDEB_Saisie.Range("O4").value)
            rs.Fields("Type").value = wshDEB_Saisie.Range("F4").value
            rs.Fields("Beneficiaire").value = wshDEB_Saisie.Range("F6").value
            rs.Fields("Reference").value = wshDEB_Saisie.Range("M6").value
            rs.Fields("Total").value = wshDEB_Saisie.Range("O6").value
            rs.Fields("No_Compte").value = wshDEB_Saisie.Range("Q" & l).value
            rs.Fields("Compte").value = wshDEB_Saisie.Range("E" & l).value
            rs.Fields("Total").value = wshDEB_Saisie.Range("H" & l).value
            rs.Fields("CodeTaxe").value = wshDEB_Saisie.Range("I" & l).value
            rs.Fields("TPS").value = wshDEB_Saisie.Range("J" & l).value
            rs.Fields("TVQ").value = wshDEB_Saisie.Range("K" & l).value
            rs.Fields("Crédit_TPS").value = wshDEB_Saisie.Range("L" & l).value
            rs.Fields("Crédit_TVQ").value = wshDEB_Saisie.Range("M" & l).value
            rs.Fields("AutreRemarque").value = ""
            rs.Fields("TimeStamp").value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        rs.update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("DEB_Trans_Add_Record_To_DB()", timerStart)

End Sub

Sub DEB_Trans_Add_Record_Locally(r As Long) 'Write records locally
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentDebTransNo As Long
    currentDebTransNo = wshDEB_Saisie.Range("B1").value
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshDébours_Trans.Range("A99999").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Integer
    For i = 9 To r
        wshDébours_Trans.Range("A" & rowToBeUsed).value = currentDebTransNo
        wshDébours_Trans.Range("B" & rowToBeUsed).value = CDate(wshDEB_Saisie.Range("O4").value)
        wshDébours_Trans.Range("C" & rowToBeUsed).value = wshDEB_Saisie.Range("F4").value
        wshDébours_Trans.Range("D" & rowToBeUsed).value = wshDEB_Saisie.Range("F6").value
        wshDébours_Trans.Range("E" & rowToBeUsed).value = wshDEB_Saisie.Range("M6").value
        wshDébours_Trans.Range("F" & rowToBeUsed).value = wshDEB_Saisie.Range("Q" & i).value
        wshDébours_Trans.Range("G" & rowToBeUsed).value = wshDEB_Saisie.Range("E" & i).value
        wshDébours_Trans.Range("H" & rowToBeUsed).value = wshDEB_Saisie.Range("H" & i).value
        wshDébours_Trans.Range("I" & rowToBeUsed).value = wshDEB_Saisie.Range("I" & i).value
        wshDébours_Trans.Range("J" & rowToBeUsed).value = wshDEB_Saisie.Range("J" & i).value
        wshDébours_Trans.Range("K" & rowToBeUsed).value = wshDEB_Saisie.Range("K" & i).value
        wshDébours_Trans.Range("L" & rowToBeUsed).value = wshDEB_Saisie.Range("L" & i).value
        wshDébours_Trans.Range("M" & rowToBeUsed).value = wshDEB_Saisie.Range("M" & i).value
        wshDébours_Trans.Range("N" & rowToBeUsed).value = ""
        wshDébours_Trans.Range("O" & rowToBeUsed).value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call Output_Timer_Results("DEB_Trans_Add_Record_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub DEB_Saisie_GL_Posting_Preparation() '2024-06-05 @ 18:28

    Dim timerStart As Double: timerStart = Timer

    Dim montant As Double
    Dim dateDebours As Date
    Dim descGL_Trans As String, source As String
    
    dateDebours = wshDEB_Saisie.Range("O4").value
    descGL_Trans = wshDEB_Saisie.Range("F4").value & " - " & _
                   wshDEB_Saisie.Range("F6").value & " [" & _
                   wshDEB_Saisie.Range("M6").value & "]"
    source = "DÉBOURS-" & Format(wshDEB_Saisie.Range("B1").value, "000000")
    
    Dim myArray() As String
    ReDim myArray(1 To 16, 1 To 4)
    
    'Disbursement Total (wshDEB_Saisie.Range("O6"))
    montant = wshDEB_Saisie.Range("O6").value
    If montant Then
        myArray(1, 1) = "1000"
        myArray(1, 2) = "Encaisse"
        myArray(1, 3) = -montant
        myArray(1, 4) = ""
    End If
    
    'Process every lines
    Dim lastUsedRow As Long
    lastUsedRow = wshDEB_Saisie.Range("E99").End(xlUp).row

    Dim l As Long, arrRow As Long
    arrRow = 2 '1 is already used
    For l = 9 To lastUsedRow
        myArray(arrRow, 1) = wshDEB_Saisie.Range("Q" & l).value
        myArray(arrRow, 2) = wshDEB_Saisie.Range("E" & l).value
        myArray(arrRow, 3) = wshDEB_Saisie.Range("N" & l).value
        myArray(arrRow, 4) = ""
        arrRow = arrRow + 1
        
        If wshDEB_Saisie.Range("L" & l).value <> 0 Then
            myArray(arrRow, 1) = "1200"
            myArray(arrRow, 2) = "TPS payée"
            myArray(arrRow, 3) = wshDEB_Saisie.Range("L" & l).value
            myArray(arrRow, 4) = ""
            arrRow = arrRow + 1
        End If

        If wshDEB_Saisie.Range("M" & l).value <> 0 Then
            myArray(arrRow, 1) = "1201"
            myArray(arrRow, 2) = "TVQ payée"
            myArray(arrRow, 3) = wshDEB_Saisie.Range("M" & l).value
            myArray(arrRow, 4) = ""
            arrRow = arrRow + 1
        End If
    Next l
   
    Call FAC_Finale_GL_Posting_To_DB(dateDebours, descGL_Trans, source, myArray)
    Call FAC_Finale_GL_Posting_Locally(dateDebours, descGL_Trans, source, myArray)
    
    Call Output_Timer_Results("FAC_Finale_GL_Posting_Preparation()", timerStart)

End Sub

Public Sub DEB_Saisie_Clear_All_Cells()

    'Vide les cellules
    Application.EnableEvents = False
    With wshDEB_Saisie
        .Range("F4:H4, F6:L6, O6, M6, E9:O23, Q9:Q23").Clearcontents
        .Range("O4").value = Format(Now(), "dd/mm/yyyy")
        .Range("F4").Select
        .Range("F4").Activate
    End With
    Application.EnableEvents = True

End Sub

