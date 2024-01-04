Attribute VB_Name = "modJE"
Option Explicit

Sub JE_Post()

    If IsDateValide = False Then Exit Sub
    
    If IsEcritureBalance = False Then Exit Sub
    
    Dim rowEJLast As Long
    rowEJLast = wshJE.Range("D99").End(xlUp).row  'Last Used Row in wshJE
    If IsEcritureValide(rowEJLast) = False Then Exit Sub
    
    Dim rowGLTrans, rowGLTransFirst As Long
    'Détermine la prochaine ligne disponible
    rowGLTrans = wshGLFACTrans.Range("C99999").End(xlUp).row + 1  'First Empty Row in wshGL
    rowGLTransFirst = rowGLTrans
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    Call AddGLTransRecordToDB(rowEJLast)

'    'Les lignes subséquentes sont en police blanche...
'    With wshGL.Range("D" & (rowGLTransFirst + 1) & ":F" & (rowGLTrans - 1)).Font
'        .Color = vbWhite
'    End With
    
'    'Ajoute des bordures à l'entrée de journal (extérieur)
'    Dim r1 As Range
'    Set r1 = wshGL.Range("D" & rowGLTransFirst & ":K" & (rowGLTrans - 2))
'    r1.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=vbBlack
'
'    With wshGL.Range("H" & (rowGLTrans - 2) & ":K" & (rowGLTrans - 2))
'        .Font.Italic = True
'        .Font.Bold = True
'        With .Interior
'            .Pattern = xlSolid
'            .PatternColorIndex = xlAutomatic
'            .ThemeColor = xlThemeColorDark1
'            .TintAndShade = -0.149998474074526
'            .PatternTintAndShade = 0
'        End With
'        .Borders(xlInsideVertical).LineStyle = xlNone
'    End With
    
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
    rowEJAuto = wshJERecurrente.Range("D99999").End(xlUp).row + 3 'First available Row in wshJERecurrente
    rowEJAutoSave = rowEJAuto
    
    Dim r As Integer
    For r = 9 To ll
        wshJERecurrente.Range("C" & rowEJAuto).value = EJAutoNo
        wshJERecurrente.Range("D" & rowEJAuto).value = wshJE.Range("K" & r).value
        wshJERecurrente.Range("E" & rowEJAuto).value = wshJE.Range("D" & r).value
        wshJERecurrente.Range("F" & rowEJAuto).value = wshJE.Range("G" & r).value
        wshJERecurrente.Range("G" & rowEJAuto).value = wshJE.Range("H" & r).value
        wshJERecurrente.Range("H" & rowEJAuto).value = wshJE.Range("I" & r).value
        wshJERecurrente.Range("I" & rowEJAuto).value = "=ROW()"
        rowEJAuto = rowEJAuto + 1
    Next
    'Ligne de description
    wshJERecurrente.Range("C" & rowEJAuto).value = EJAutoNo
    wshJERecurrente.Range("E" & rowEJAuto).value = wshJE.Range("E6").value
    wshJERecurrente.Range("I" & rowEJAuto).value = "=ROW()"
    rowEJAuto = rowEJAuto + 1
    'Ligne vide
    wshJERecurrente.Range("C" & rowEJAuto).value = EJAutoNo
    wshJERecurrente.Range("I" & rowEJAuto).value = "=ROW()"
    rowEJAuto = rowEJAuto + 1
    
    'Ajoute la description dans la liste des E/J automatiques (K1:L99999)
    Dim rowEJAutoDesc As Long
    rowEJAutoDesc = wshJERecurrente.Range("K99999").End(xlUp).row + 1 'First available Row in wshJERecurrente
    wshJERecurrente.Range("K" & rowEJAutoDesc).value = wshJE.Range("E6").value
    wshJERecurrente.Range("L" & rowEJAutoDesc).value = EJAutoNo

    'Ajoute des bordures à l'entrée de journal récurrente
    Dim r1 As Range
    Set r1 = wshJERecurrente.Range("D" & rowEJAutoSave & ":H" & (rowEJAuto - 2))
    r1.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=vbBlack

End Sub

Sub LoadJEAutoIntoJE(EJAutoDesc As String, NoEJAuto As Long)

    'On copie l'E/J automatique vers wshEJ
    Dim rowJEAuto, rowJE As Long
    rowJEAuto = wshJERecurrente.Range("C99999").End(xlUp).row  'Last Row used in wshJERecuurente
    
    Call wshJEClearAllCells
    rowJE = 9
    
    Dim r As Long
    For r = 2 To rowJEAuto
        If wshJERecurrente.Range("C" & r).value = NoEJAuto And wshJERecurrente.Range("D" & r).value <> "" Then
            wshJE.Range("D" & rowJE).value = wshJERecurrente.Range("E" & r).value
            wshJE.Range("G" & rowJE).value = wshJERecurrente.Range("F" & r).value
            wshJE.Range("H" & rowJE).value = wshJERecurrente.Range("G" & r).value
            wshJE.Range("I" & rowJE).value = wshJERecurrente.Range("H" & r).value
            wshJE.Range("K" & rowJE).value = wshJERecurrente.Range("D" & r).value
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
        .Range("D9:F22,G9:G22,H9:H22,I9:J22,K9:K22").ClearContents
        .ckbRecurrente = False
    End With

End Sub

Sub BuildDate(cell As String, r As Range)
        Dim d, m, y As Integer
        Dim strDateConsruite As String
        Dim dateValide As Boolean
        dateValide = True

        cell = Replace(cell, "/", "")
        cell = Replace(cell, "-", "")

        'Utilisation de la date du jour pour valuer par défaut
        d = Day(Now())
        m = Month(Now())
        y = Year(Now())

        Select Case Len(cell)
            Case 0
                strDateConsruite = Format(d, "00") & "/" & Format(m, "00") & "/" & Format(y, "0000")
            Case 1, 2
                strDateConsruite = Format(cell, "00") & "/" & Format(m, "00") & "/" & Format(y, "0000")
            Case 3
                strDateConsruite = Format(Left(cell, 1), "00") & "/" & Format(Mid(cell, 2, 2), "00") & "/" & Format(y, "0000")
            Case 4
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(y, "0000")
            Case 6
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & "20" & Format(Mid(cell, 5, 2), "00")
            Case 8
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(Mid(cell, 5, 4), "0000")
            Case Else
                dateValide = False
        End Select
        dateValide = IsDate(strDateConsruite)

    If dateValide Then
        r.value = Format(strDateConsruite, "dd/mm/yyyy")
    Else
        MsgBox "La saisie est invalide...", vbInformation, "Il est impossible de construire une date"
    End If

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

    IsEcritureValide = False
    If rmax < 10 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        IsEcritureValide = False
    Else
        IsEcritureValide = True
    End If

End Function

Sub AddGLTransRecordToDB(r As Long) 'Write/Update a record to external .xlsx file
    Dim FullFileName As String
    Dim SheetName As String
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim maxEJNo As Long, lastRow As Long, nextJENo As Long
    
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
        lastRow = 1
    Else
        lastRow = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new ID
    nextJENo = lastRow + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & SheetName & "$] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    
    For l = 9 To r + 2
        rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("No_EJ").value = nextJENo
        rs.Fields("Date").value = CDate(wshJE.Range("J4").value)
        rs.Fields("Numéro Écriture").value = nextJENo
        rs.Fields("Source").value = wshJE.Range("E4").value
        If l <= r Then
            rs.Fields("No_Compte").value = wshJE.Range("K" & l).value
            rs.Fields("Compte").value = wshJE.Range("D" & l).value
            rs.Fields("Débit").value = wshJE.Range("G" & l).value
            rs.Fields("Crédit").value = wshJE.Range("H" & l).value
            rs.Fields("AutreRemarque").value = wshJE.Range("I" & l).value
        Else
            If l = r + 1 Then
                rs.Fields("Compte").value = wshJE.Range("E6").value
            End If
        End If
        rs.Fields("No.Ligne").Formula = "=LIGNE()"
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

