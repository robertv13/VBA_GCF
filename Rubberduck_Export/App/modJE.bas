Attribute VB_Name = "modJE"
Option Explicit

Sub JE_Post()

    If IsDateValide = False Then Exit Sub
    
    If IsEcritureBalance = False Then Exit Sub
    
    Dim RowEJLast As Long
    RowEJLast = wshJE.Range("D99").End(xlUp).row  'Last Used Row in wshJE
    If IsEcritureValide(RowEJLast) = False Then Exit Sub
    
    Dim rowGLTrans, rowGLTransFirst As Long
    'Détermine la prochaine ligne disponible
    rowGLTrans = wshGL.Range("C99999").End(xlUp).row + 1  'First Empty Row in wshGL
    rowGLTransFirst = rowGLTrans
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    FromJE2GL RowEJLast, rowGLTrans

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
        SaveEJRecurrente RowEJLast
    End If
    
    With wshJE
        'Increment Next JE number
        .Range("B1").value = .Range("B1").value + 1
        Call wshJEClearAllCells
        .Range("E4").Activate
    End With
    
End Sub

Sub FromJE2GL(rEJLast As Long, ByRef rGLTrans)

    Dim l As Long
    With wshGL
        For l = 9 To rEJLast + 2
            .Range("C" & rGLTrans).value = wshJE.Range("B1").value
            .Range("D" & rGLTrans).value = Format(CDate(wshJE.Range("J4").value), "dd/mm/yyyy")
            .Range("E" & rGLTrans).value = wshJE.Range("B1").value
            .Range("F" & rGLTrans).value = wshJE.Range("E4").value
            If l <= rEJLast Then
                .Range("G" & rGLTrans).value = wshJE.Range("K" & l).value
                .Range("H" & rGLTrans).value = wshJE.Range("D" & l).value
                .Range("I" & rGLTrans).value = wshJE.Range("G" & l).value
                .Range("J" & rGLTrans).value = wshJE.Range("H" & l).value
                .Range("K" & rGLTrans).value = wshJE.Range("I" & l).value
            Else
                If l = rEJLast + 1 Then
                    .Range("H" & rGLTrans).value = wshJE.Range("E6").value
                End If
            End If
            .Range("L" & rGLTrans).value = "=ROW()"
            rGLTrans = rGLTrans + 1
        Next l
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
        Dim strDateJour, strDateConsruite As String
        Dim dateValide As Boolean
        dateValide = True

        cell = Replace(cell, "/", "")
        cell = Replace(cell, "-", "")

        'Utilisation de la date du jour
        strDateJour = Now
        d = Day(strDateJour)
        m = Month(strDateJour)
        y = Year(strDateJour)

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

Sub GLTrans_Import() '2024-01-01 @ 09:30
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshGLFACTrans.Range("C1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GLTrans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-01-01 @ 08:30
                     
    'Set up source and destination ranges
    Dim sourceRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets("GLTrans").UsedRange

    Dim destinationRange As Range
    Set destinationRange = wshGLFACTrans.Range("C1")

    'Copy data, using Range to Range
    sourceRange.Copy destinationRange
    wshGLFACTrans.Range("C1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wshGLFACTrans.Range("C999999").End(xlUp).row
    
    With wshGLFACTrans
        With .Range("C2" & ":L" & lastRow)
            .HorizontalAlignment = xlCenter
        End With
        With .Range("F2:F" & lastRow & ", H2:H" & lastRow & ", K2:K" & lastRow)
            .HorizontalAlignment = xlLeft
        End With
        With .Range("I2:I" & lastRow & ", J2:J" & lastRow)
            .HorizontalAlignment = xlRight
        End With
        .Range("I2:I" & lastRow & ", J2:J" & lastRow).NumberFormat = "#,##0.00 $"
        .Range("D2:D" & lastRow).NumberFormat = "dd/mm/yyyy"
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Sub AddGLTransToDB(r As Long) 'Write/Update a record to external .xlsx file
    Dim FullFileName As String
    Dim SheetName As String
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim maxID As Long
    Dim lastRow As Long
    Dim nextID As Long
    
    Application.ScreenUpdating = False
    
    FullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    SheetName = "GLTrans"
    
    'Initialize connection, connection string & open the connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Set rs = CreateObject("ADODB.Recordset")

    If r < 0 Then 'Soft delete
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & SheetName & "$] WHERE TEC_ID=" & Abs(r), conn, 2, 3
        If Not rs.EOF Then
            'Update the "IsDeleted" field to mark the record as deleted
            rs.Fields("DateSaisie").value = Now
            rs.Fields("EstDetruit").value = True
            rs.Fields("VersionApp").value = gAppVersion
            rs.Update
        Else
            ' Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut être trouvé!", vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    Else
        'If r is 0, add a new record; otherwise, update an existing record
        If r = 0 Then 'Add a record
        'SQL select command to find the next available ID
            strSQL = "SELECT MAX(TEC_ID) AS MaxID FROM [" & SheetName & "$]"
        
            'Open recordset to find out the MaxID
            rs.Open strSQL, conn
            
            'Get the last used row
            If IsNull(rs.Fields("MaxID").value) Then
                ' Handle empty table (assign a default value, e.g., 1)
                lastRow = 1
            Else
                lastRow = rs.Fields("MaxID").value
            End If
            
            'Calculate the new ID
            nextID = lastRow + 1
        
            'Close the previous recordset, no longer needed and open an empty recordset
            rs.Close
            rs.Open "SELECT * FROM [" & SheetName & "$] WHERE 1=0", conn, 2, 3
            rs.AddNew
            
            'Add fields to the recordset before updating it
            rs.Fields("TEC_ID").value = nextID
            rs.Fields("Prof_ID").value = wshAdmin.Range("TEC_Prof_ID")
            rs.Fields("Prof").value = frmSaisieHeures.cmbProfessionnel.value
            rs.Fields("Date").value = CDate(frmSaisieHeures.txtDate.value)
            rs.Fields("Client_ID").value = wshAdmin.Range("TEC_Client_ID")
            rs.Fields("ClientNom").value = frmSaisieHeures.txtClient.value
            rs.Fields("Description").value = frmSaisieHeures.txtActivite.value
            rs.Fields("Heures").value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
            rs.Fields("CommentaireNote").value = frmSaisieHeures.txtCommNote.value
            rs.Fields("EstFacturable").value = frmSaisieHeures.chbFacturable.value
            rs.Fields("DateSaisie").value = Now
            rs.Fields("EstFacturee").value = False
            rs.Fields("DateFacturee").value = ""
            rs.Fields("EstDetruit").value = False
            rs.Fields("VersionApp").value = gAppVersion
            rs.Fields("NoFacture").value = ""
        Else 'Update an existing record
            'Open the recordset for the specified ID
            rs.Open "SELECT * FROM [" & SheetName & "$] WHERE TEC_ID=" & r, conn, 2, 3
            If Not rs.EOF Then
                'Update fields for the existing record
                rs.Fields("Client_ID").value = wshAdmin.Range("TEC_Client_ID")
                rs.Fields("ClientNom").value = frmSaisieHeures.txtClient.value
                rs.Fields("Description").value = frmSaisieHeures.txtActivite.value
                rs.Fields("Heures").value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
                rs.Fields("CommentaireNote").value = frmSaisieHeures.txtCommNote.value
                rs.Fields("EstFacturable").value = frmSaisieHeures.chbFacturable.value
                rs.Fields("DateSaisie").value = Now
                rs.Fields("VersionApp").value = gAppVersion
            Else
                'Handle the case where the specified ID is not found
                MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut être trouvé!", vbExclamation
                rs.Close
                conn.Close
                Exit Sub
            End If
        End If
    End If

    'Update the recordset (create the record)
    rs.Update
    rs.Close
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

End Sub
