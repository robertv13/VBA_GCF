Attribute VB_Name = "modGL_Posting"
Option Explicit

Sub GL_Posting_To_DB(df, desc, source, arr As Variant, ByRef glEntryNo) 'Generic routine 2024-06-06 @ 07:00

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Posting:GL_Posting_To_DB", 0)

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans"
    
    'Initialize connection, connection string, open the connection and declare rs Object
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJNo As Long
    strSQL = "SELECT MAX(No_Entr�e) AS MaxEJNo FROM [" & destinationTab & "$]"

    'Open recordset to find out the next JE number
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 0
    Else
        lastJE = rs.Fields("MaxEJNo").Value
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
                rs.Fields("No_Entr�e") = glEntryNo
                rs.Fields("Date") = CDate(df)
                rs.Fields("Description") = desc
                rs.Fields("Source") = source
                rs.Fields("No_Compte") = arr(i, 1)
                rs.Fields("Compte") = arr(i, 2)
                If arr(i, 3) > 0 Then
                    rs.Fields("D�bit") = arr(i, 3)
                Else
                    rs.Fields("Cr�dit") = -arr(i, 3)
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

    'Lib�rer la m�moire
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
                .Range("A" & rowToBeUsed).Value = glEntryNo
                .Range("B" & rowToBeUsed).Value = CDate(df)
                .Range("C" & rowToBeUsed).Value = desc
                .Range("D" & rowToBeUsed).Value = source
                .Range("E" & rowToBeUsed).Value = arr(i, 1)
                .Range("F" & rowToBeUsed).Value = arr(i, 2)
                If arr(i, 3) > 0 Then
                     .Range("G" & rowToBeUsed).Value = CDbl(arr(i, 3))
                Else
                     .Range("H" & rowToBeUsed).Value = -CDbl(arr(i, 3))
                End If
                .Range("I" & rowToBeUsed).Value = arr(i, 4)
                .Range("J" & rowToBeUsed).Value = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
                rowToBeUsed = rowToBeUsed + 1
            End If
        Next i
    End With
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modGL_Posting:GL_Posting_Locally", startTime)

End Sub


