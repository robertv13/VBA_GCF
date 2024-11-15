Attribute VB_Name = "modGL_Stuff"
Option Explicit

Public Sub Get_GL_Trans_With_AF(glCode As String, dateDeb As Date, dateFin As Date) '2024-11-08 @ 09:34

    Dim ws As Worksheet: Set ws = wshGL_Trans
    
    'Où allons-nous mettre les résultats ?
    Dim rngResult As Range
    Set rngResult = ws.Range("P1").CurrentRegion.Offset(1, 0)
    rngResult.ClearContents
    Set rngResult = ws.Range("P1").CurrentRegion
    
    'Où sont les données à traiter ?
    Dim rngSource As Range
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    'Rien à traiter
    If lastUsedRow < 2 Then
        Exit Sub
    End If
    Set rngSource = ws.Range("A1:J" & lastUsedRow)
    
    'Quels sont les critères ?
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("L2:N3")
    With ws
        .Range("L3").value = glCode
        .Range("M3").value = ">=" & CLng(dateDeb)
        .Range("N3").value = "<=" & CLng(dateFin)
    End With
    
    'On documente le processus
    ws.Range("M6:M10").ClearContents
    ws.Range("M6").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    ws.Range("M7").value = rngSource.Address
    ws.Range("M8").value = rngCriteria.Address
    ws.Range("M9").value = rngResult.Address
    
    'Go, on execute le AdvancedFilter
    rngSource.AdvancedFilter xlFilterCopy, _
                             rngCriteria, _
                             rngResult, _
                             False
    
    'Combien y a-t-il de transactions dans le résultat ?
    lastUsedRow = ws.Cells(ws.rows.count, "P").End(xlUp).row
    ws.Range("M10").value = lastUsedRow
    Set rngResult = ws.Range("P1:Y" & lastUsedRow)

    If lastUsedRow > 2 Then
        With ws.Sort
            .SortFields.Clear
                .SortFields.Add _
                    key:=ws.Range("Q2"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par date de transaction
                .SortFields.Add _
                    key:=ws.Range("P2"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par numéro d'écriture
            .SetRange rngResult
            .Header = xlYes
            .Apply
        End With
    End If

End Sub

Sub GL_Posting_To_DB(df, desc, source, arr As Variant, ByRef glEntryNo) 'Generic routine 2024-06-06 @ 07:00

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Posting:GL_Posting_To_DB", 0)

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans"
    
    'Initialize connection, connection string, open the connection and declare rs Object
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJNo As Long
    strSQL = "SELECT MAX(No_Entrée) AS MaxEJNo FROM [" & destinationTab & "$]"

    'Open recordset to find out the next JE number
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 0
    Else
        lastJE = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new JE number
    glEntryNo = lastJE + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    Dim TimeStamp As String
    Dim i As Long, j As Long
    'Loop through the array and post each row
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) = "" Then GoTo Nothing_to_Post
            rs.AddNew
                rs.Fields("No_Entrée") = glEntryNo
                rs.Fields("Date") = CDate(df)
                rs.Fields("Description") = desc
                rs.Fields("Source") = source
                rs.Fields("No_Compte") = arr(i, 1)
                rs.Fields("Compte") = arr(i, 2)
                If arr(i, 3) > 0 Then
                    rs.Fields("Débit") = arr(i, 3)
                Else
                    rs.Fields("Crédit") = -arr(i, 3)
                End If
                rs.Fields("AutreRemarque") = arr(i, 4)
                TimeStamp = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
                rs.Fields("TimeStamp") = TimeStamp
                Debug.Print "GL_Trans - " & CDate(Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
            rs.update
Nothing_to_Post:
    Next i

    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modGL_Posting:GL_Posting_To_DB", startTime)

End Sub

Sub GL_Posting_Locally(df, desc, source, arr As Variant, ByRef glEntryNo) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Posting:GL_Posting_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in GL_Trans ?
    Dim rowToBeUsed As Long
    rowToBeUsed = wshGL_Trans.Range("A99999").End(xlUp).row + 1
    
    Dim i As Long, j As Long
    'Loop through the array and post each row
    With wshGL_Trans
        For i = LBound(arr, 1) To UBound(arr, 1)
            If arr(i, 1) <> "" Then
                .Range("A" & rowToBeUsed).value = glEntryNo
                .Range("B" & rowToBeUsed).value = CDate(df)
                .Range("C" & rowToBeUsed).value = desc
                .Range("D" & rowToBeUsed).value = source
                .Range("E" & rowToBeUsed).value = arr(i, 1)
                .Range("F" & rowToBeUsed).value = arr(i, 2)
                If arr(i, 3) > 0 Then
                     .Range("G" & rowToBeUsed).value = CDbl(arr(i, 3))
                Else
                     .Range("H" & rowToBeUsed).value = -CDbl(arr(i, 3))
                End If
                .Range("I" & rowToBeUsed).value = arr(i, 4)
                .Range("J" & rowToBeUsed).value = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
                rowToBeUsed = rowToBeUsed + 1
            End If
        Next i
    End With
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modGL_Posting:GL_Posting_Locally", startTime)

End Sub
