Attribute VB_Name = "modGL_Stuff"
Option Explicit

Public Sub GL_Get_Account_Trans_AF(glNo As String, dateDeb As Date, dateFin As Date, ByRef rResult As Range)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport:GL_Get_Account_Trans_AF(" & glNo & " - De " & dateDeb & " � " & dateFin & ")", 0)

    'Les donn�es � AF proviennent de GL_Trans
    Dim ws As Worksheet: Set ws = wshGL_Trans
    
    'wshGL_Trans_AF#1

    'Effacer les donn�es de la derni�re utilisation
    ws.Range("M6:M10").ClearContents
    ws.Range("M6").Value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'D�finir le range pour la source des donn�es en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_GL_Trans[#All]")
    ws.Range("M7").Value = rngData.Address
    
    'D�finir le range des crit�res
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("L2:N3")
    ws.Range("L3").Value = glNo
    ws.Range("M3").Value = ">=" & CLng(dateDeb)
    ws.Range("N3").Value = "<=" & CLng(dateFin)
    ws.Range("M8").Value = rngCriteria.Address
    
    'D�finir le range des r�sultats et effacer avant le traitement
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
        
    'Quels sont les r�sultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "P").End(xlUp).row
    ws.Range("M10").Value = lastUsedRow - 1 & " lignes"
    
    If lastUsedRow > 2 Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=wshGL_Trans.Range("T2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Tri par num�ro de compte
            .SortFields.Add key:=wshGL_Trans.Range("Q2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Tri par date
            .SortFields.Add key:=wshGL_Trans.Range("P2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Tri par num�ro d'�criture
            .SetRange wshGL_Trans.Range("P2:Y" & lastUsedRow) 'Set Range
            .Apply 'Apply Sort
        End With
    End If

    'Retourne le Range des r�sultats
    Set rResult = wshGL_Trans.Range("P1:Y" & lastUsedRow)
    
    'Lib�rer la m�moire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modGL_Rapport:GL_Get_Account_Trans_AF", startTime)

End Sub

Sub GL_Posting_To_DB(df, desc, source, arr As Variant, ByRef GLEntryNo) 'Generic routine 2024-06-06 @ 07:00

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Posting:GL_Posting_To_DB", 0)

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans$"
    
    'Initialize connection, connection string, open the connection and declare rs Object
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJNo As Long
    strSQL = "SELECT MAX(NoEntr�e) AS MaxEJNo FROM [" & destinationTab & "]"

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

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim timeStamp As String
    Dim i As Long, j As Long
    'Loop through the array and post each row
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) = "" Then GoTo Nothing_to_Post
            rs.AddNew
                'RecordSet are ZERO base, and Enums are not, so the '-1' is mandatory !!!
                rs.Fields(fGlTNoEntr�e - 1).Value = GLEntryNo
                rs.Fields(fGlTDate - 1).Value = CDate(df)
                rs.Fields(fGlTDescription - 1).Value = desc
                rs.Fields(fGlTSource - 1).Value = source
                rs.Fields(fGlTNoCompte - 1).Value = arr(i, 1)
                rs.Fields(fGlTCompte - 1).Value = arr(i, 2)
                If arr(i, 3) > 0 Then
                    rs.Fields(fGlTD�bit - 1).Value = arr(i, 3)
                Else
                    rs.Fields(fGlTCr�dit - 1).Value = -arr(i, 3)
                End If
                rs.Fields(fGlTAutreRemarque - 1).Value = arr(i, 4)
                timeStamp = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
                rs.Fields(fGlTTimeStamp - 1).Value = timeStamp
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

Sub GL_Posting_Locally(df, desc, source, arr As Variant, ByRef GLEntryNo) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Posting:GL_Posting_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in GL_Trans ?
    Dim rowToBeUsed As Long
    rowToBeUsed = wshGL_Trans.Cells(wshGL_Trans.Rows.count, 1).End(xlUp).row + 1
    
    Dim i As Long, j As Long
    'Loop through the array and post each row
    With wshGL_Trans
        For i = LBound(arr, 1) To UBound(arr, 1)
            If arr(i, 1) <> "" Then
                .Range("A" & rowToBeUsed).Value = GLEntryNo
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

Sub GL_BV_Ajouter_Shape_Retour()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Stuff:GL_BV_Ajouter_Shape_Retour", 0)
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Dim btn As Shape
    Dim leftPosition As Double
    Dim topPosition As Double

    'Trouver la derni�re ligne de la plage L4:T*
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.count, "M").End(xlUp).row

    'Calculer les positions (Left & Top) du bouton
    leftPosition = ws.Range("T" & LastRow).Left
    topPosition = ws.Range("S" & LastRow).Top + (2 * ws.Range("S" & LastRow).Height)

    ' Ajouter une Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, Left:=leftPosition, Top:=topPosition, _
                                                    Width:=90, Height:=30)
    With btn
        .Name = "shpRetour"
        .TextFrame2.TextRange.Text = "Retour"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.size = 14
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.ForeColor.RGB = RGB(166, 166, 166)
        .OnAction = "GL_BV_Effacer_Zone_Et_Shape"
    End With
    
    'Lib�rer la m�moire
    Set btn = Nothing
    Set ws = Nothing
    
    Call Log_Record("modGL_Stuff:GL_BV_Ajouter_Shape_Retour", startTime)

End Sub

Sub GL_BV_Effacer_Zone_Et_Shape()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Stuff:GL_BV_Effacer_Zone_Et_Shape", 0)
    
    'Effacer la plage
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Application.EnableEvents = False
    ws.Range("L1:T" & ws.Cells(ws.Rows.count, "M").End(xlUp).row).offset(3, 0).Clear
    Application.EnableEvents = True

    'Trouver la derni�re ligne de la plage M4:T*
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.count, "M").End(xlUp).row

    'Supprimer la shape
    On Error Resume Next
    ws.Shapes("shpRetour").Delete
    On Error GoTo 0

    Call GL_BV_Hide_Dynamic_Shape
    
    'Ramener le focus � C4
    Application.EnableEvents = False
    ws.Range("C4").Select
    Application.EnableEvents = True
    
    'Lib�rer la m�moire
    Set ws = Nothing
    
    Call Log_Record("modGL_Stuff:GL_BV_Effacer_Zone_Et_Shape", startTime)

End Sub
