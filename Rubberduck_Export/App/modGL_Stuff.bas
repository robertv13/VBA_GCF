Attribute VB_Name = "modGL_Stuff"
Option Explicit

'Structure pour une écriture comptable (données communes)
Public Type clsGL_Entry '2025-06-08 @ 06:59

    DateTrans As Date
    Source As String
    NoCompte As String
    AutreRemarque As String
    
End Type

'Structure pour une écriture comptable (données spécifiques à chaque ligne)
Public Type clsGL_EntryLine '2025-06-08 @ 07:02

    NoCompte As String
    description As String
    Montant As Double
    
End Type

Public Sub GL_Get_Account_Trans_AF(glNo As String, dateDeb As Date, dateFin As Date, ByRef rResult As Range)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Stuff:GL_Get_Account_Trans_AF", glNo & " - De " & dateDeb & " à " & dateFin, 0)

    'Les données à AF proviennent de GL_Trans
    Dim ws As Worksheet: Set ws = wsdGL_Trans
    
    'wsdGL_Trans_AF#1

    'Effacer les données de la dernière utilisation
    ws.Range("M6:M10").ClearContents
    ws.Range("M6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_GL_Trans[#All]")
    ws.Range("M7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("L2:N3")
    ws.Range("L3").Value = glNo
    ws.Range("M3").Value = ">=" & CLng(dateDeb)
    ws.Range("N3").Value = "<=" & CLng(dateFin)
    ws.Range("M8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("P1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("P1:Y1")
    ws.Range("M9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "P").End(xlUp).Row
    ws.Range("M10").Value = lastUsedRow - 1 & " lignes"
    
    If lastUsedRow > 2 Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsdGL_Trans.Range("T2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Tri par numéro de compte
            .SortFields.Add key:=wsdGL_Trans.Range("Q2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Tri par date
            .SortFields.Add key:=wsdGL_Trans.Range("P2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Tri par numéro d'écriture
            .SetRange wsdGL_Trans.Range("P2:Y" & lastUsedRow) 'Set Range
            .Apply 'Apply Sort
        End With
    End If

    'Retourne le Range des résultats
    Set rResult = wsdGL_Trans.Range("P1:Y" & lastUsedRow)
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modGL_Stuff:GL_Get_Account_Trans_AF", vbNullString, startTime)

End Sub

Sub GL_Posting_To_DB(df As Date, desc As String, Source As String, arr As Variant, ByRef GLEntryNo As Long) 'Generic routine 2024-06-06 @ 07:00

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Stuff:GL_Posting_To_DB", vbNullString, 0)

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans$"
    
    'Initialize connection, connection string, open the connection and declare rs Object
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJNo As Long
    strSQL = "SELECT MAX(NoEntrée) AS MaxEJNo FROM [" & destinationTab & "]"

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
    GLEntryNo = lastJE + 1

    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim i As Long, j As Long
    'Loop through the array and post each row
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) = vbNullString Then GoTo Nothing_to_Post
            rs.AddNew
                'RecordSet are ZERO base, and Enums are not, so the '-1' is mandatory !!!
                rs.Fields(fGlTNoEntrée - 1).Value = GLEntryNo
                rs.Fields(fGlTDate - 1).Value = CDate(df)
                rs.Fields(fGlTDescription - 1).Value = desc
                rs.Fields(fGlTSource - 1).Value = Source
                rs.Fields(fGlTNoCompte - 1).Value = arr(i, 1)
                rs.Fields(fGlTCompte - 1).Value = arr(i, 2)
                If arr(i, 3) > 0 Then
                    rs.Fields(fGlTDébit - 1).Value = arr(i, 3)
                Else
                    rs.Fields(fGlTCrédit - 1).Value = -arr(i, 3)
                End If
                rs.Fields(fGlTAutreRemarque - 1).Value = arr(i, 4)
                rs.Fields(fGlTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rs.Update
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
    
    Call Log_Record("modGL_Stuff:GL_Posting_To_DB", vbNullString, startTime)

End Sub

Sub GL_Posting_Locally(df As Date, desc As String, Source As String, arr As Variant, ByRef GLEntryNo As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("*** modGL_Stuff:GL_Posting_Locally", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in GL_Trans ?
    Dim rowToBeUsed As Long
    rowToBeUsed = wsdGL_Trans.Cells(wsdGL_Trans.Rows.count, 1).End(xlUp).Row + 1
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim i As Long, j As Long
    'Loop through the array and post each row
    With wsdGL_Trans
        For i = LBound(arr, 1) To UBound(arr, 1)
            If arr(i, 1) <> vbNullString Then
                .Range("A" & rowToBeUsed).Value = GLEntryNo
                .Range("B" & rowToBeUsed).Value = CDate(df)
                .Range("C" & rowToBeUsed).Value = desc
                .Range("D" & rowToBeUsed).Value = Source
                .Range("E" & rowToBeUsed).Value = arr(i, 1)
                .Range("F" & rowToBeUsed).Value = arr(i, 2)
                If arr(i, 3) > 0 Then
                     .Range("G" & rowToBeUsed).Value = CDbl(arr(i, 3))
                Else
                     .Range("H" & rowToBeUsed).Value = -CDbl(arr(i, 3))
                End If
                .Range("I" & rowToBeUsed).Value = arr(i, 4)
                .Range("J" & rowToBeUsed).Value = Format$(timeStamp, "dd/mm/yyyy hh:mm:ss")
                rowToBeUsed = rowToBeUsed + 1
                Call Log_Record("   modGL_Stuff:GL_Posting_Locally", -1)
            End If
        Next i
    End With
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modGL_Stuff:GL_Posting_Locally", vbNullString, startTime)

End Sub

Sub GL_BV_Ajouter_Shape_Retour()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Stuff:GL_BV_Ajouter_Shape_Retour", vbNullString, 0)
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Dim btn As Shape
    Dim leftPosition As Double
    Dim topPosition As Double

    'Trouver la dernière ligne de la plage L4:T*
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "M").End(xlUp).Row

    If lastRow >= 5 Then
        'Calculer les positions (Left & Top) du bouton
        leftPosition = ws.Range("T" & lastRow).Left
        topPosition = ws.Range("S" & lastRow).Top + (2 * ws.Range("S" & lastRow).Height)
    
        ' Ajouter une Shape
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, Left:=leftPosition, Top:=topPosition, _
                                                        Width:=90, Height:=30)
        With btn
            .Name = "shpRetour"
            .TextFrame2.TextRange.text = "Retour"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame2.TextRange.Font.size = 14
            .TextFrame2.TextRange.Font.Bold = True
            .TextFrame2.HorizontalAnchor = msoAnchorCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .Fill.ForeColor.RGB = RGB(166, 166, 166)
            .OnAction = "GL_BV_Effacer_Zone_Et_Shape"
        End With
    End If
    
    'Libérer la mémoire
    Set btn = Nothing
    Set ws = Nothing
    
    Call Log_Record("modGL_Stuff:GL_BV_Ajouter_Shape_Retour", vbNullString, startTime)

End Sub

Sub GL_BV_Effacer_Zone_Et_Shape()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Stuff:GL_BV_Effacer_Zone_Et_Shape", vbNullString, 0)
    
    'Effacer la plage
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Application.EnableEvents = False
    ws.Range("L1:T" & ws.Cells(ws.Rows.count, "M").End(xlUp).Row).Offset(3, 0).Clear
    Application.EnableEvents = True

    'Supprimer les formes shpRetour
    Call GL_BV_SupprimerToutesLesFormes_shpRetour(ws)

    Call GL_BV_Hide_Dynamic_Shape
    
    'Ramener le focus à C4
    Application.EnableEvents = False
    ws.Range("D4").Select
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modGL_Stuff:GL_BV_Effacer_Zone_Et_Shape", vbNullString, startTime)

End Sub

Sub GL_BV_EffacerZoneBV(w As Worksheet)

    Application.EnableEvents = False
    Dim lastUsedRow As Long
    lastUsedRow = w.Cells(w.Rows.count, "D").End(xlUp).Row
    If lastUsedRow >= 4 Then
        w.Range("D4:G" & lastUsedRow).Clear
    End If
    Application.EnableEvents = True

End Sub

Sub GL_BV_EffacerZoneTransactionsDetaillees(w As Worksheet)

    Application.EnableEvents = False
    Dim lastUsedRow As Long
    lastUsedRow = w.Cells(w.Rows.count, "M").End(xlUp).Row
    
    Application.EnableEvents = False
    If lastUsedRow >= 4 Then
        w.Range("L4:T" & lastUsedRow).Clear
    End If
    Application.EnableEvents = True
    
    'Supprimer les formes 'shpRetour'
    Call GL_BV_SupprimerToutesLesFormes_shpRetour(w)

    Application.EnableEvents = True

End Sub

Sub GL_BV_SupprimerToutesLesFormes_shpRetour(w As Worksheet)

    Dim shp As Shape

    For Each shp In w.Shapes
        If shp.Name = "shpRetour" Then
            shp.Delete
        End If
    Next shp
    
End Sub


