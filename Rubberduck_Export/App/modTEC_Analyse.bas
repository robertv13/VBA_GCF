Attribute VB_Name = "modTEC_Analyse"
Option Explicit

Sub TEC_Sort_Group_And_Subtotal()

    Application.ScreenUpdating = False
    
    Dim lastUsedRow As Long
    Dim firstEmptyCol As Long
    
    'Set the source worksheet, lastUsedRow and lastUsedCol
    Dim wsSource As Worksheet: Set wsSource = wshTEC_Local
    'Find the last row with data in the source worksheet
    lastUsedRow = wsSource.Cells(wsSource.rows.count, "A").End(xlUp).row
    'Find the first empty column from the left in the source worksheet
    firstEmptyCol = 1
    Do Until IsEmpty(wsSource.Cells(2, firstEmptyCol))
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    'Set the current worksheet as the result
    Dim wsDest As Worksheet: Set wsDest = wshTEC_Analyse
    'Remove existing subtotals in the destination worksheet
    wsDest.Cells.RemoveSubtotal
    
    Dim destLastUsedRow As Long
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "B").End(xlUp).row
    wsDest.Range("A6:H" & destLastUsedRow).ClearContents
    
    Dim i As Long, r As Long
    r = 6
    Application.EnableEvents = False
    For i = 3 To lastUsedRow
        'Conditions for exclusion (adjust as needed)
        If wsSource.Cells(i, 14).value <> "VRAI" And _
            wsSource.Cells(i, 12).value <> "VRAI" And _
            wsSource.Cells(i, 10).value = "VRAI" Then
                wsDest.Cells(r, 1).value = wsSource.Cells(i, ftecTEC_ID).value
                wsDest.Cells(r, 2).value = wsSource.Cells(i, ftecProf_ID).value
                wsDest.Cells(r, 3).value = wsSource.Cells(i, ftecClientNom).value
                wsDest.Cells(r, 5).value = wsSource.Cells(i, ftecDate).value
                wsDest.Cells(r, 6).value = wsSource.Cells(i, ftecProf).value
                wsDest.Cells(r, 7).value = wsSource.Cells(i, ftecDescription).value
                wsDest.Cells(r, 8).value = wsSource.Cells(i, ftecHeures).value
                wsDest.Cells(r, 9).value = wsSource.Cells(i, ftecCommentaireNote).value
                r = r + 1
        End If
    Next i
    Application.EnableEvents = False
   
    'Find the last row in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row

    'Sort by Client_ID (column E) and Date (column D) in the destination worksheet
    wsDest.Sort.SortFields.clear
    wsDest.Sort.SortFields.add key:=wsDest.Range("C6:C" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.add key:=wsDest.Range("E6:E" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.add key:=wsDest.Range("B6:B" & destLastUsedRow), Order:=xlAscending
    
    With wsDest.Sort
        .SetRange wsDest.Range("A5:H" & destLastUsedRow)
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Add subtotals for hours (column G) at each change in ClientNom_ID (column B) in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row
    Application.DisplayAlerts = False
    wsDest.Range("A6:I" & destLastUsedRow).Subtotal GroupBy:=3, Function:=xlSum, _
        TotalList:=Array(8), Replace:=True, PageBreaks:=False, SummaryBelowData:=False
    Application.DisplayAlerts = True
    wsDest.Range("A:B").EntireColumn.Hidden = True

    'Group the data to show subtotals in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row
    wsDest.Outline.ShowLevels ROwLevels:=2
    
    'Add a formula to sum the billed amounts
    wshTEC_Analyse.Range("D6").formula = "=SUM(D7:D" & destLastUsedRow & ")"
    
    'Change the format of 'Total general' row
    With wshTEC_Analyse.Range("D6")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Font
            .Color = -16776961
            .TintAndShade = 0
            .Bold = True
            .Size = 12
        End With
    End With
    With wshTEC_Analyse.Range("H6")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .Color = -16776961
            .TintAndShade = 0
            .Bold = True
            .Size = 12
        End With
    End With
    
    'Change the format of Group Totals rows
    For r = 7 To destLastUsedRow
        If wsDest.Range("A" & r).value = "" Then
            With wsDest.Range("C" & r).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With wsDest.Range("C" & r).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            With wsDest.Range("H" & r).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With wsDest.Range("H" & r).Font
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If
    Next r
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub Build_Hours_Summary(client As String, r As Long)

    If r < 7 Then Exit Sub
    
    Dim lastUsedRow As Long
    lastUsedRow = wshTEC_Analyse.Cells(wshTEC_Analyse.rows.count, "A").End(xlUp).row
    wshTEC_Analyse.Range("K:Q").clear
    Call Delete_CheckBox
    
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
    Dim i As Long, saveR As Long
    r = r + 1 'Summary starts on the next line (first line of expanded lines)
    saveR = r
    i = r
    Do Until Cells(i, 5) = ""
        If Cells(i, 6).value <> "" Then
'            t = t + Cells(i, 7).value
            If dictHours.Exists(Cells(i, 6).value) Then
                dictHours(Cells(i, 6).value) = dictHours(Cells(i, 6).value) + Cells(i, 8).value
            Else
                dictHours.add Cells(i, 6).value, Cells(i, 8).value
            End If
        End If
        i = i + 1
    Loop

    Dim Prof As Variant
    Dim ProfID As Integer
    wshTEC_Analyse.Range("Q" & r).value = 0 'Reset the total WIP value
    For Each Prof In Fn_Sort_Dictionary_By_Value(dictHours, True) ' Sort dictionary by hours in descending order
        Cells(r, 11).value = Prof
        Dim strProf As String
        strProf = Prof
        ProfID = Fn_GetID_From_Initials(strProf)
        Cells(r, 12).HorizontalAlignment = xlRight
        Cells(r, 12).NumberFormat = "#,##0.00"
        Cells(r, 12).value = dictHours(Prof)
        Dim tauxHoraire As Currency
        tauxHoraire = Fn_Get_Hourly_Rate(ProfID, "2024-07-15")
        Cells(r, 13).value = tauxHoraire
        Cells(r, 14).NumberFormat = "#,##0.00$"
        Cells(r, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
        Cells(r, 14).HorizontalAlignment = xlRight
        r = r + 1
    Next Prof
    
    'Sort the summary by rate (descending value) if required
    If r - 1 > saveR Then
        Dim rngSort As Range
        Set rngSort = wshTEC_Analyse.Range(wshTEC_Analyse.Cells(saveR, 11), wshTEC_Analyse.Cells(r - 1, 14))
        rngSort.Sort Key1:=wshTEC_Analyse.Cells(saveR, 13), Order1:=xlDescending, header:=xlNo
    End If
    
    'Add totals to the summary
    Dim rTotal As Long
    rTotal = r
    With Cells(rTotal, 12)
        .HorizontalAlignment = xlRight
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
'        .value = Format(t, "#,##0.00")
        .Font.Bold = True
    End With
    
    With Cells(r, 14)
        .HorizontalAlignment = xlRight
'        .value = Format(tdollars, "#,##0.00$")
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
        .Font.Bold = True
    End With
    
    With Range("K" & saveR & ":N" & r).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    With Range("L" & r & ", N" & r)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With

    'Save the TOTAL WIP value
    With wshTEC_Analyse.Range("P" & saveR)
        .value = "Valeur TEC:"
        .Font.Italic = True
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With wshTEC_Analyse.Range("Q" & saveR)
        .NumberFormat = "#,##0.00$"
        .value = wshTEC_Analyse.Range("N" & r).value
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    'Create a visual clue is amounts are different
    With wshTEC_Analyse.Range("Q" & r)
        Dim formula As String
        formula = "=IF(N" & r & "<>Q" & saveR & ", N" & r & "-Q" & saveR & ",""""" & ")"
        .formula = formula
        .NumberFormat = "#,##0.00$"
    End With
    
    Call Add_And_Modify_Checkbox(saveR, r)
    
    'Clean up - 2024-07-11 @ 15:20
    Set dictHours = Nothing
    Set rngSort = Nothing
    
End Sub
    
Sub FAC_Projets_Détails_Add_Record_To_DB(ClientID As Long, fr As Long, lr As Long, ByRef ProjetID As Long) 'Write a record to MASTER.xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modTEC_ANalyse:FAC_Projet_Détails_Add_Record_To_DB()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "FAC_Projets_Détails"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    
    'First SQL - SQL query to find the maximum value in the first column
    Dim strSQL As String
    strSQL = "SELECT MAX(ProjetID) AS MaxValue FROM [" & destinationTab & "$]"
    rs.Open strSQL, conn

    'Get the maximum value
    Dim MaxValue As Long
    If IsNull(rs.Fields("MaxValue").value) Then
        'Handle empty table (assign a default value, e.g., 1)
        ProjetID = 1
    Else
        ProjetID = rs.Fields("MaxValue").value + 1
    End If
    Debug.Print "ProjetID = " & ProjetID
    
    'Close the previous recordset (no longer needed)
    rs.Close
    
    'Second SQL - SQL query to add the new records
    strSQL = "SELECT * FROM [" & destinationTab & "$] WHERE 1=0"
    rs.Open strSQL, conn, 2, 3
    
    'Read all line from TEC_Analyse
    Dim dateTEC As String, TimeStamp As String
    Dim l As Long
    For l = fr To lr
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("ProjetID").value = ProjetID
            rs.Fields("NomClient").value = wshTEC_Analyse.Range("C" & l).value
            rs.Fields("ClientID").value = ClientID
            rs.Fields("TECID").value = wshTEC_Analyse.Range("A" & l).value
            rs.Fields("ProfID").value = wshTEC_Analyse.Range("B" & l).value
            dateTEC = Format(wshTEC_Analyse.Range("E" & l).value, "dd/mm/yyyy")
            rs.Fields("Date").value = dateTEC
            rs.Fields("Prof").value = wshTEC_Analyse.Range("F" & l).value
            rs.Fields("estDétruite").value = False
            rs.Fields("Heures").value = CDbl(wshTEC_Analyse.Range("H" & l).value)
            TimeStamp = Format(Now(), "dd/mm/yyyy hh:mm:ss")
            rs.Fields("TimeStamp").value = TimeStamp
        rs.update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Output_Timer_Results("modTEC_ANalyse:FAC_Projet_Détails_Add_Record_To_DB()", timerStart)

End Sub

Sub FAC_Projets_Détails_Add_Record_Locally(ClientID As Long, fr As Long, lr As Long, ProjetID As Long) 'Write records locally
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modTEC_Analyse:FAC_Projet_Détails_Add_Record_Locally()")
    
    Application.ScreenUpdating = False
    
    'What is the last used row in FAC_Projets_Détails?
    Dim lastUsedRow As Long, rn As Long
    lastUsedRow = wshFAC_Projets_Détails.Range("A99999").End(xlUp).row
    rn = lastUsedRow + 1
    
    Dim dateTEC As String, TimeStamp As String
    Dim i As Integer
    For i = fr To lr
        wshFAC_Projets_Détails.Range("A" & rn).value = ProjetID
        wshFAC_Projets_Détails.Range("B" & rn).value = wshTEC_Analyse.Range("C" & i).value
        wshFAC_Projets_Détails.Range("C" & rn).value = ClientID
        wshFAC_Projets_Détails.Range("D" & rn).value = wshTEC_Analyse.Range("A" & i).value
        wshFAC_Projets_Détails.Range("E" & rn).value = wshTEC_Analyse.Range("B" & i).value
        dateTEC = Format(wshTEC_Analyse.Range("E" & i).value, "dd/mm/yyyy")
        wshFAC_Projets_Détails.Range("F" & rn).value = dateTEC
        wshFAC_Projets_Détails.Range("G" & rn).value = wshTEC_Analyse.Range("F" & i).value
        wshFAC_Projets_Détails.Range("H" & rn).value = wshTEC_Analyse.Range("H" & i).value
        wshFAC_Projets_Détails.Range("I" & rn).value = False
        TimeStamp = Format(Now(), "dd/mm/yyyy hh:mm:ss")
        wshFAC_Projets_Détails.Range("J" & rn).value = TimeStamp
        rn = rn + 1
    Next i
    
    Call Output_Timer_Results("modTEC_Analyse:FAC_Projet_Détails_Add_Record_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub Soft_Delete_If_Value_Is_Found_In_Master_Details(filePath As String, _
                                                    sheetName As String, _
                                                    columnName As String, _
                                                    valueToFind As Variant) '2024-07-19 @ 15:31
    'Create a new ADODB connection
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    'Open the connection to the closed workbook
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    
    'Update the rows to mark as deleted (soft delete)
    Dim strSQL As String
    strSQL = "UPDATE [" & sheetName & "$] SET estDétruite = True WHERE [" & columnName & "] = '" & Replace(valueToFind, "'", "''") & "'"
    cn.Execute strSQL
    
    'Close the connection
    cn.Close
    Set cn = Nothing
    
End Sub

Sub FAC_Projets_Entête_Add_Record_To_DB(ProjetID As Long, _
                                        NomClient As String, _
                                        ClientID As Long, _
                                        dte As String, _
                                        hono As Double, _
                                        ByRef arr As Variant) 'Write a record to MASTER.xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modTEC_ANalyse:FAC_Projet_Entête_Add_Record_To_DB()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "FAC_Projets_Entête"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "$] WHERE 1=0"
    
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conn, 2, 3
    
    Dim TimeStamp As String
    Dim c As Integer
    Dim l As Long
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("ProjetID").value = ProjetID
        rs.Fields("NomClient").value = NomClient
        rs.Fields("ClientID").value = ClientID
        rs.Fields("Date").value = dte
        rs.Fields("HonoTotal").value = hono
        For c = 1 To UBound(arr, 1)
            rs.Fields("Prof" & c).value = arr(c, 1)
            rs.Fields("Hres" & c).value = arr(c, 2)
            rs.Fields("TauxH" & c).value = arr(c, 3)
            rs.Fields("Hono" & c).value = arr(c, 4)
        Next c
        rs.Fields("estDétruite").value = False
        TimeStamp = Format(Now(), "dd/mm/yyyy hh:mm:ss")
        rs.Fields("TimeStamp").value = TimeStamp
    rs.update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Output_Timer_Results("modTEC_ANalyse:FAC_Projet_Entête_Add_Record_To_DB()", timerStart)

End Sub

Sub FAC_Projets_Entête_Add_Record_Locally(ProjetID As Long, NomClient As String, ClientID As Long, dte As String, hono As Double, ByRef arr As Variant) 'Write records locally
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modTEC_Analyse:FAC_Projet_Entête_Add_Record_Locally()")
    
    Application.ScreenUpdating = False
    
    'What is the last used row in FAC_Projets_Détails?
    Dim lastUsedRow As Long, rn As Long
    lastUsedRow = wshFAC_Projets_Entête.Range("A99999").End(xlUp).row
    rn = lastUsedRow + 1
    
    Dim dateTEC As String, TimeStamp As String
    wshFAC_Projets_Entête.Range("A" & rn).value = ProjetID
    wshFAC_Projets_Entête.Range("B" & rn).value = NomClient
    wshFAC_Projets_Entête.Range("C" & rn).value = ClientID
    wshFAC_Projets_Entête.Range("D" & rn).value = dte
    wshFAC_Projets_Entête.Range("E" & rn).value = hono
    'Assign values from the array to the worksheet using .Cells
    Dim i As Integer, j As Integer
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            wshFAC_Projets_Entête.Cells(rn, 6 + (i - 1) * UBound(arr, 2) + j - 1).value = arr(i, j)
        Next j
    Next i
    wshFAC_Projets_Entête.Range("Z" & rn).value = False
    TimeStamp = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    wshFAC_Projets_Entête.Range("AA" & rn).value = TimeStamp
    
    Call Output_Timer_Results("modTEC_Analyse:FAC_Projet_Entête_Add_Record_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub Soft_Delete_If_Value_Is_Found_In_Master_Entete(filePath As String, _
                                                   sheetName As String, _
                                                   columnName As String, _
                                                   valueToFind As Variant) '2024-07-19 @ 15:31
    'Create a new ADODB connection
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    'Open the connection to the closed workbook
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    
    'Update the rows to mark as deleted (soft delete)
    Dim strSQL As String
    strSQL = "UPDATE [" & sheetName & "$] SET estDétruite = True WHERE [" & columnName & "] = '" & Replace(valueToFind, "'", "''") & "'"
    cn.Execute strSQL
    
    'Close the connection
    cn.Close
    Set cn = Nothing
    
End Sub

Sub Add_And_Modify_Checkbox(startRow As Long, lastRow As Long)
    
    'Set your worksheet (adjust this to match your worksheet name)
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Define the range for the summary
    Dim summaryRange As Range
    Set summaryRange = ws.Range(ws.Cells(startRow, 11), ws.Cells(lastRow, 14)) 'Columns K to N
    
    'Add an ActiveX checkbox next to the summary in column O
    Dim checkBox As OLEObject
    With ws
        Set checkBox = .OLEObjects.add(ClassType:="Forms.CheckBox.1", _
                    Left:=.Cells(lastRow, 15).Left + 5, _
                    Top:=.Cells(lastRow, 15).Top, width:=80, Height:=16)
        
        'Modify checkbox properties
        With checkBox.Object
            .Caption = "On facture"
            .Font.Size = 11  ' Set font size
            .Font.Bold = True  ' Set font bold
            .ForeColor = RGB(0, 0, 255)  ' Set font color (Blue)
            .BackColor = RGB(200, 255, 200)  ' Set background color (Light Green)
        End With
    End With
    
End Sub

Sub Delete_CheckBox()

    'Set your worksheet (adjust this to match your worksheet name)
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Check if CheckBox1 exists and then delete it
    Dim checkBox As OLEObject
    Dim i As Integer
    For i = 1 To 5
        On Error Resume Next
        Set checkBox = ws.OLEObjects("CheckBox" & i)
        If Not checkBox Is Nothing Then
            checkBox.delete
        End If
        On Error GoTo 0
    Next i
End Sub

Sub Groups_SubTotals_Collapse_A_Client(r As Long)

    'Set the worksheet you want to work on
    Dim ws As Worksheet: Set ws = wshTEC_Analyse

    'Loop through each row starting at row r
    Dim saveR As Long
    saveR = r
    Do While wshTEC_Analyse.Range("A" & r).value <> ""
        r = r + 1
    Loop

    r = r - 1
    ws.rows.ClearOutline
    
    ws.rows(saveR & ":" & r).Group
    ws.rows(saveR & ":" & r).Hidden = True
    
End Sub

