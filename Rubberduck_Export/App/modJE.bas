Attribute VB_Name = "modJE"
Option Explicit

Sub JE_Post()

    If IsDateValide = False Then Exit Sub
    
    If IsEcritureBalance = False Then Exit Sub
    
    Dim rowEJLast As Long
    rowEJLast = wshJE.Range("D99").End(xlUp).row  'Last Used Row in wshJE
    If IsEcritureValide(rowEJLast) = False Then Exit Sub
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    Call AddGLTransRecordToDB(rowEJLast)
    
    If wshJE.ckbRecurrente = True Then
        SaveEJRecurrente rowEJLast
    End If
    
    With wshJE
        'Increment Next JE number
        .Range("B1").value = .Range("B1").value + 1
        Call wshJEClearAllCells
        .Range("E4").Activate
    End With
    
End Sub

Sub SaveEJRecurrente(ll As Long)

    Dim EJAutoNo As Long
    EJAutoNo = wshJERecurrente.Range("B1").value
    wshJERecurrente.Range("B1").value = wshJERecurrente.Range("B1").value + 1
    
    Dim rowEJAuto, rowEJAutoSave As Long
    rowEJAuto = wshJERecurrente.Range("C99999").End(xlUp).row + 2 'First available Row in wshJERecurrente
    rowEJAutoSave = rowEJAuto
    
    Dim r As Integer
    For r = 9 To ll
        wshJERecurrente.Range("C" & rowEJAuto).value = EJAutoNo
        wshJERecurrente.Range("D" & rowEJAuto).value = wshJE.Range("E6").value
        wshJERecurrente.Range("E" & rowEJAuto).value = wshJE.Range("K" & r).value
        wshJERecurrente.Range("F" & rowEJAuto).value = wshJE.Range("D" & r).value
        wshJERecurrente.Range("G" & rowEJAuto).value = wshJE.Range("G" & r).value
        wshJERecurrente.Range("H" & rowEJAuto).value = wshJE.Range("H" & r).value
        wshJERecurrente.Range("I" & rowEJAuto).value = wshJE.Range("I" & r).value
        wshJERecurrente.Range("J" & rowEJAuto).value = "=ROW()"
        rowEJAuto = rowEJAuto + 1
    Next
    'Ligne vide
    wshJERecurrente.Range("C" & rowEJAuto).value = EJAutoNo
    wshJERecurrente.Range("J" & rowEJAuto).value = "=ROW()"
    rowEJAuto = rowEJAuto + 1
    
    'Ajoute la description dans la liste des E/J automatiques (K1:L99999)
    Dim rowEJAutoDesc As Long
    rowEJAutoDesc = wshJERecurrente.Range("L99999").End(xlUp).row + 1 'First available Row in wshJERecurrente
    wshJERecurrente.Range("L" & rowEJAutoDesc).value = wshJE.Range("E6").value
    wshJERecurrente.Range("M" & rowEJAutoDesc).value = EJAutoNo

    'r1.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=vbBlack

End Sub

Sub LoadJEAutoIntoJE(EJAutoDesc As String, NoEJAuto As Long)

    'On copie l'E/J automatique vers wshEJ
    Dim rowJEAuto, rowJE As Long
    rowJEAuto = wshEJRecurrente.Range("C99999").End(xlUp).row  'Last Row used in wshJERecuurente
    
    Call wshJEClearAllCells
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

Sub wshJEClearAllCells()

    'Efface toutes les cellules de la feuille
    With wshJE
        .Range("E4,J4,E6:J6").ClearContents
        .Range("D9:F23,G9:G23,H9:H23,I9:J23,K9:K23").ClearContents
        .ckbRecurrente = False
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

Sub AddGLTransRecordToDB(r As Long) 'Write/Update a record to external .xlsx file
    Dim FullFileName As String
    Dim SheetName As String
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim maxEJNo As Long, lastJE As Long, nextJENo As Long
    
    Application.ScreenUpdating = False
    
    FullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    SheetName = "GLTrans"
    
    'Initialize connection, connection string & open the connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    strSQL = "SELECT MAX(No_EJ) AS MaxEJNo FROM [" & SheetName & "$]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    If IsNull(rs.Fields("MaxEJNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new ID
    nextJENo = lastJE + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & SheetName & "$] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    
    For l = 9 To r + 1
        rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("No_EJ").value = nextJENo
        rs.Fields("Date").value = CDate(wshJE.Range("J4").value)
        rs.Fields("Numéro Écriture").value = nextJENo
        rs.Fields("Source").value = wshJE.Range("E4").value
        rs.Fields("Description").value = wshJE.Range("E6").value
        rs.Fields("No_Compte").value = wshJE.Range("K" & l).value
        rs.Fields("Compte").value = wshJE.Range("D" & l).value
        rs.Fields("Débit").value = wshJE.Range("G" & l).value
        rs.Fields("Crédit").value = wshJE.Range("H" & l).value
        rs.Fields("AutreRemarque").value = wshJE.Range("I" & l).value
        'rs.Fields("No.Ligne").Formula = "=ROW()"
        rs.Update
    Next l
    
    rs.Close
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

