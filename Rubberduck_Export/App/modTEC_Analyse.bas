Attribute VB_Name = "modTEC_Analyse"
Option Explicit

Sub TEC_Sort_Group_And_Subtotal() '2024-08-24 @ 08:10

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Analyse:TEC_Sort_Group_And_Subtotal", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsDest As Worksheet: Set wsDest = wshTEC_Analyse
    
    'Remove existing subtotals in the destination worksheet
    wsDest.Cells.RemoveSubtotal
'    Call Log_Record("     modTEC_Analyse:TEC_Sort_Group_And_Subtotal - Les anciens SubTotal ont été effacés", -1)
    
    'Clear the worksheet from row 6 until the last row used
    Dim destLastUsedRow As Long
    destLastUsedRow = wsDest.Cells(wsDest.Rows.count, "B").End(xlUp).row
    If destLastUsedRow < 6 Then destLastUsedRow = 6
    wsDest.Range("A6:I" & destLastUsedRow).Clear
    
    'Créer un dict pour tous les clients FACTURABLES
    Dim wsClientsMF As Worksheet: Set wsClientsMF = wshBD_Clients
    Dim lastUsedRowClient
    lastUsedRowClient = wsClientsMF.Cells(wsClientsMF.Rows.count, "B").End(xlUp).row
    Dim dictClients As Dictionary
    Set dictClients = New Dictionary
    Dim clientData As Variant
    'Charger toutes les données des clients dans un tableau
    clientData = wsClientsMF.Range(wsClientsMF.Cells(2, fClntFMClientID), _
                               wsClientsMF.Cells(lastUsedRowClient, fClntFMClientNom)).value

    ' Parcourir le tableau pour ajouter les clients facturables au dictionnaire
    Dim i As Long
    For i = 1 To UBound(clientData, 1)
        If Fn_Is_Client_Facturable(clientData(i, 2)) = True Then
            dictClients.Add CStr(clientData(i, 2)), clientData(i, 1)
        End If
    Next i

    Dim lastUsedRow As Long, firstEmptyCol As Long
    
    'Set the source worksheet, lastUsedRow and lastUsedCol
    Dim wsSource As Worksheet: Set wsSource = wshTEC_Local
    'Find the last row with data in the source worksheet
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
    'Find the first empty column from the left in the source worksheet
    firstEmptyCol = 1
    Do Until IsEmpty(wsSource.Cells(2, firstEmptyCol))
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    Application.EnableEvents = False
    
    'Appel à AdvancedFilter # 2 dans TEC_Local
    Call Get_TEC_For_Client_AF("", CLng(CDate(wsDest.Range("H3").value)), "VRAI", "FAUX", "FAUX")
    
    Dim lastUsedResult As Long
    lastUsedResult = wshTEC_Local.Cells(wshTEC_Local.Rows.count, "AQ").End(xlUp).row
    
    'Charger les données sources dans un tableau (beaucoup plus rapide)
    Dim sourceData As Variant
    Dim rowCount As Long
    sourceData = wsSource.Range("AQ3:AX" & lastUsedResult).value
    rowCount = UBound(sourceData, 1)

    'Initialiser un tableau pour les données de sortie (beaucoup plus rapide)
    Dim outputData() As Variant
    ReDim outputData(1 To rowCount, 1 To 8)
    
    Dim r As Long: r = 1
    Dim codeClient As String, nomClientFromMF As String
    
    For i = 1 To rowCount
        'Vérifier la condition d'exclusion
        If dictClients.Exists(sourceData(i, 5)) Then
            codeClient = sourceData(i, 5)
            nomClientFromMF = dictClients(codeClient)
            'Ajouter les données au tableau de sortie
            outputData(r, 1) = sourceData(i, 1)
            outputData(r, 2) = sourceData(i, 2)
            outputData(r, 3) = nomClientFromMF
            outputData(r, 5) = sourceData(i, 4)
            outputData(r, 6) = sourceData(i, 3)
            outputData(r, 7) = sourceData(i, 7)
            outputData(r, 8) = sourceData(i, 8)
            r = r + 1
        End If
    Next i
    
    'Écrire les données de sortie dans la feuille & formater quelques colonnes
    If r > 1 Then
        wsDest.Range("A7:H" & r - 1 + 6).value = outputData
        'Formats
        wsDest.Range("E7:F" & r - 1 + 6).HorizontalAlignment = xlCenter
        wsDest.Range("H7:H" & r - 1 + 6).NumberFormat = "#,##0.00"
    End If
    
    Application.EnableEvents = False
   
    'Find the last row in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row

    'Sort by ClientID (column E) and Date (column D) in the destination worksheet
    wsDest.Sort.SortFields.Clear
    wsDest.Sort.SortFields.Add key:=wsDest.Range("C7:C" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.Add key:=wsDest.Range("E7:E" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.Add key:=wsDest.Range("B7:B" & destLastUsedRow), Order:=xlAscending
    
    With wsDest.Sort
        .SetRange wsDest.Range("A7:I" & destLastUsedRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Add subtotals for hours (column H) at each change in nomClientMF (column C) in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row
    Application.DisplayAlerts = False
    wsDest.Range("A6:H" & destLastUsedRow).Subtotal GroupBy:=3, Function:=xlSum, _
            TotalList:=Array(8), Replace:=True, PageBreaks:=False, SummaryBelowData:=False
    Application.DisplayAlerts = True
    wsDest.Range("A:B").EntireColumn.Hidden = True
'    Call Log_Record("     modTEC_Analyse:TEC_Sort_Group_And_Subtotal - Le GroupBy est complété", -1)

    'Group the data to show subtotals in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row
    wsDest.Outline.ShowLevels RowLevels:=2
'    Call Log_Record("     modTEC_Analyse:TEC_Sort_Group_And_Subtotal - Le 'ShowLevels est ajusté à 2", -1)
    
    'Add a formula to sum the billed amounts at the top row
    wsDest.Range("D7").formula = "=SUM(D8:D" & destLastUsedRow & ")"
    wsDest.Range("D7").NumberFormat = "#,##0.00 $"
    
    'Change the format of the top row (Total General)
    With wsDest.Range("C7:D7")
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
            .size = 12
        End With
    End With
    
    'Change the format of the top row (Hours)
    With wsDest.Range("H7")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .Name = "Aptos Narrow"
            .Bold = True
            .size = 12
        End With
    End With
    
    'Change the format of all Client's Total rows
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
            With wsDest.Range("C" & r)
'                If InStr(.value, "Total ") = 1 Then
'                    .value = Mid(.value, 7)
'                End If
                If .value = "Total général" Then
                    .value = "G r a n d   T o t a l"
                End If
            End With
            'Mettre de l'emphase sur les cellules d'heures, si le montant du projet <> 0,00 $
            If wsDest.Range("D" & r).value = 0 Then
                With wsDest.Range("H" & r).Font
                    .Name = "Aptos Narrow"
                    .size = 12
                    .Bold = True
                End With
            End If
        End If
    Next r
    
    'Set conditional formats for total hours (Client's total)
    Dim rngTotals As Range: Set rngTotals = wsDest.Range("C8:C" & destLastUsedRow)
    Call Apply_Conditional_Formatting_Alternate_On_Column_H(rngTotals, destLastUsedRow)
    
    'Bring in all the invoice requests
    Call Bring_In_Existing_Invoice_Requests(destLastUsedRow)
    
    'Clean up the summary area of the worksheet
    Call Clean_Up_Summary_Area(wsDest)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set dictClients = Nothing
    Set rngTotals = Nothing
    Set wsClientsMF = Nothing
    Set wsDest = Nothing
    Set wsSource = Nothing
    
    Call Log_Record("modTEC_Analyse:TEC_Sort_Group_And_Subtotal", "", startTime)

End Sub

Sub Clean_Up_Summary_Area(ws As Worksheet)

    Application.EnableEvents = False
    
    'Clean up the summary area (columns K to Q)
    ws.Range("J:P").Clear
    'Erase any checkbox left over
    Call Delete_CheckBox
    
    Application.EnableEvents = True

End Sub

Sub Apply_Conditional_Formatting_Alternate_On_Column_H(rng As Range, lastUsedRow As Long)

    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Loop each cell in column C to find Totals row
    Dim totalRange As Range, cell As Range
    For Each cell In rng
        If InStr(1, cell.value, "Total ", vbTextCompare) > 0 Then
            If totalRange Is Nothing Then
                Set totalRange = ws.Cells(cell.row, 8) 'Column H
            Else
                Set totalRange = Union(totalRange, ws.Cells(cell.row, 8))
            End If
        End If
    Next cell
    
    'Check if any total rows were found
    rng.FormatConditions.Delete

    'Define conditional formatting rules for the total rows
    If Not totalRange Is Nothing Then
        'Clear existing conditional formatting rules in the totalRange
        totalRange.FormatConditions.Delete
        
        'Define conditional formatting rules for the totalRange
        With totalRange.FormatConditions
    
            'Rule for values > 50 (Highest priority)
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="50"
            With .item(.count)
                .Interior.Color = RGB(255, 0, 0) 'Red color
            End With
    
            'Rule for values > 25
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="25"
            With .item(.count)
                .Interior.Color = RGB(255, 165, 0) 'Orange color
            End With
    
            'Rule for values > 10
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="10"
            With .item(.count)
                .Interior.Color = RGB(255, 255, 0) 'Yellow color
            End With
    
            'Rule for values > 5
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="5"
            With .item(.count)
                .Interior.Color = RGB(144, 238, 144) 'Light green color
            End With
        End With
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set totalRange = Nothing
    Set ws = Nothing
            
End Sub

Sub Build_Hours_Summary(rowSelected As Long)

    If rowSelected < 8 Then Exit Sub
    
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Determine the last row used
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    'Clear the Hours Summary area
    Call Clean_Up_Summary_Area(ws)
    
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
    Dim i As Long, saveR As Long
    rowSelected = rowSelected + 1 'Summary starts on the next line (first line of expanded lines)
    saveR = rowSelected
    i = rowSelected
    Do Until Cells(i, 5) = ""
        If Cells(i, 6).value <> "" Then
            If dictHours.Exists(Cells(i, 6).value) Then
                dictHours(Cells(i, 6).value) = dictHours(Cells(i, 6).value) + Cells(i, 8).value
            Else
                dictHours.Add Cells(i, 6).value, Cells(i, 8).value
            End If
            Cells(i, 8).Font.Color = RGB(166, 166, 166) 'RMV_15
        End If
        i = i + 1
    Loop

    Dim prof As Variant
    Dim profID As Long
    Dim tauxHoraire As Currency
    
    Application.EnableEvents = False
    
    ws.Range("O" & rowSelected).value = 0 'Reset the total WIP value
    For Each prof In Fn_Sort_Dictionary_By_Value(dictHours, True) ' Sort dictionary by hours in descending order
        Cells(rowSelected, 10).value = prof
        Dim strProf As String
        strProf = prof
        profID = Fn_GetID_From_Initials(strProf)
        Cells(rowSelected, "K").HorizontalAlignment = xlRight
        Cells(rowSelected, "K").NumberFormat = "#,##0.00"
        Cells(rowSelected, "K").value = dictHours(prof)
        tauxHoraire = Fn_Get_Hourly_Rate(profID, ws.Range("H3").value)
        Cells(rowSelected, "L").value = tauxHoraire
        Cells(rowSelected, "M").NumberFormat = "#,##0.00 $"
        Cells(rowSelected, "M").FormulaR1C1 = "=RC[-2]*RC[-1]"
        Cells(rowSelected, "M").HorizontalAlignment = xlRight
        rowSelected = rowSelected + 1
    Next prof
    
    'Sort the summary by rate (descending value) if required
    If rowSelected - 1 > saveR Then
        Dim rngSort As Range
        Set rngSort = ws.Range(ws.Cells(saveR, 10), ws.Cells(rowSelected - 1, 13))
        rngSort.Sort Key1:=ws.Cells(saveR, 13), Order1:=xlDescending, Header:=xlNo
    End If
    
    'Hours Total
    Dim rTotal As Long
    rTotal = rowSelected
    With Cells(rTotal, "K")
        .HorizontalAlignment = xlRight
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
'        .value = Format(t, "#,##0.00")
        .Font.Bold = True
    End With
    
    'Fees Total
    With Cells(rowSelected, "M")
        .HorizontalAlignment = xlRight
'        .value = Format(tdollars, "#,##0.00$")
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
        .Font.Bold = True
    End With
    
    With Range("J" & saveR & ":M" & rowSelected).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    With Range("K" & rowSelected & ", M" & rowSelected)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With

    'Save the TOTAL WIP value
    With ws.Range("N" & saveR)
        .value = "Valeur TEC:"
        .Font.Italic = True
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("O" & saveR)
        .NumberFormat = "#,##0.00 $"
        .value = ws.Range("M" & rowSelected).value
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    'Create a visual clue if amounts are different
    With ws.Range("O" & rowSelected)
        Dim formula As String
        formula = "=IF(M" & rowSelected & " <> O" & saveR & ", M" & rowSelected & "-O" & saveR & ",""""" & ")"
        Application.EnableEvents = False
        .formula = formula
        .NumberFormat = "#,##0.00 $"
        Application.EnableEvents = True
    End With
    
    Call Add_And_Modify_Checkbox(saveR, rowSelected)
    
    Application.EnableEvents = True

    'Libérer la mémoire
    Set dictHours = Nothing
    Set prof = Nothing
    Set rngSort = Nothing
    Set ws = Nothing
    
End Sub
    
Sub Bring_In_Existing_Invoice_Requests(activeLastUsedRow As Long)

    Dim wsSource As Worksheet: Set wsSource = wshFAC_Projets_Entête
    Dim sourceLastUsedRow As Long
    sourceLastUsedRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).row
    
    Dim wsActive As Worksheet: Set wsActive = wshTEC_Analyse
    Dim rngTotal As Range: Set rngTotal = wsActive.Range("C1:C" & activeLastUsedRow)
    
    'Analyze all Invoice Requests (one row at the time)
    
    Dim clientName As String
    Dim clientID As String
    Dim honoTotal As Double
    Dim result As Variant
    Dim i As Long, r As Long
    For i = 2 To sourceLastUsedRow
        If wsSource.Cells(i, 26).value <> "True" Then
            clientName = wsSource.Cells(i, 2).value
            clientID = wsSource.Cells(i, 3).value
            honoTotal = wsSource.Cells(i, 5).value
            'Using XLOOKUP to find the result directly
            result = Application.WorksheetFunction.XLookup("Total " & clientName, _
                                                           rngTotal, _
                                                           rngTotal, _
                                                           "Not Found", _
                                                           0, _
                                                           1)
            
            If result <> "Not Found" Then
                r = Application.WorksheetFunction.Match(result, rngTotal, 0)
                wsActive.Cells(r, 4).value = honoTotal
                wsActive.Cells(r, 4).NumberFormat = "#,##0.00 $"
            End If
        End If
    Next i

    'Libérer la mémoire
    Set rngTotal = Nothing
    Set wsActive = Nothing
    Set wsSource = Nothing
    
End Sub

Sub FAC_Projets_Détails_Add_Record_To_DB(clientID As String, fr As Long, lr As Long, ByRef projetID As Long) 'Write a record to MASTER.xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Analyse:FAC_Projets_Détails_Add_Record_To_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_Détails$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    
    'Initialize recordset
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    
    'First SQL - SQL query to find the maximum value in the first column
    Dim strSQL As String
    strSQL = "SELECT MAX(ProjetID) AS MaxValue FROM [" & destinationTab & "]"
    rs.Open strSQL, conn

    'Get the maximum value
    Dim MaxValue As Long
    If IsNull(rs.Fields("MaxValue").value) Then
        'Handle empty table (assign a default value, e.g., 1)
        projetID = 1
    Else
        projetID = rs.Fields("MaxValue").value + 1
    End If
    
    'Close the previous recordset (no longer needed)
    rs.Close
    
    'Second SQL - SQL query to add the new records
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE 1=0"
    rs.Open strSQL, conn, 2, 3
    
    'Read all line from TEC_Analyse
    Dim dateTEC As String, timeStamp As String
    Dim l As Long
    For l = fr To lr
        rs.AddNew
            'Add fields to the recordset before updating it
            'RecordSet are ZERO base, and Enums are not, so the '-1' is mandatory !!!
            rs.Fields(fFacPDProjetID - 1).value = projetID
            rs.Fields(fFacPDNomClient - 1).value = wshTEC_Analyse.Range("C" & l).value
            rs.Fields(fFacPDClientID - 1).value = clientID
            rs.Fields(fFacPDTECID - 1).value = wshTEC_Analyse.Range("A" & l).value
            rs.Fields(fFacPDProfID - 1).value = wshTEC_Analyse.Range("B" & l).value
            dateTEC = Format$(wshTEC_Analyse.Range("E" & l).value, "yyyy-mm-dd")
            rs.Fields(fFacPDDate - 1).value = dateTEC
            rs.Fields(fFacPDProf - 1).value = wshTEC_Analyse.Range("F" & l).value
            rs.Fields(fFacPDestDetruite - 1) = 0 'Faux
            rs.Fields(fFacPDHeures - 1).value = CDbl(wshTEC_Analyse.Range("H" & l).value)
            timeStamp = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            rs.Fields(fFacPDTimeStamp - 1).value = timeStamp
        rs.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modTEC_Analyse:FAC_Projets_Détails_Add_Record_To_DB", "", startTime)

End Sub

Sub FAC_Projets_Détails_Add_Record_Locally(clientID As String, fr As Long, lr As Long, projetID As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Analyse:FAC_Projets_Détails_Add_Record_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in FAC_Projets_Détails?
    Dim lastUsedRow As Long, rn As Long
    lastUsedRow = wshFAC_Projets_Détails.Cells(wshFAC_Projets_Détails.Rows.count, "A").End(xlUp).row
    rn = lastUsedRow + 1
    
    Dim dateTEC As String, timeStamp As String
    Dim i As Long
    For i = fr To lr
        wshFAC_Projets_Détails.Range("A" & rn).value = projetID
        wshFAC_Projets_Détails.Range("B" & rn).value = wshTEC_Analyse.Range("C" & i).value
        wshFAC_Projets_Détails.Range("C" & rn).value = clientID
        wshFAC_Projets_Détails.Range("D" & rn).value = wshTEC_Analyse.Range("A" & i).value
        wshFAC_Projets_Détails.Range("E" & rn).value = wshTEC_Analyse.Range("B" & i).value
        dateTEC = Format$(wshTEC_Analyse.Range("E" & i).value, "yyyy-mm-dd")
        wshFAC_Projets_Détails.Range("F" & rn).value = dateTEC
        wshFAC_Projets_Détails.Range("G" & rn).value = wshTEC_Analyse.Range("F" & i).value
        wshFAC_Projets_Détails.Range("H" & rn).value = wshTEC_Analyse.Range("H" & i).value
        wshFAC_Projets_Détails.Range("I" & rn).value = "FAUX"
        timeStamp = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        wshFAC_Projets_Détails.Range("J" & rn).value = timeStamp
        rn = rn + 1
    Next i
    
    Application.ScreenUpdating = True

    Call Log_Record("modTEC_Analyse:FAC_Projets_Détails_Add_Record_Locally", "", startTime)

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
    strSQL = "UPDATE [" & sheetName & "] SET estDetruite = -1 WHERE [" & columnName & "] = '" & Replace(valueToFind, "'", "''") & "'"
    cn.Execute strSQL
    
    'Close the connection
    cn.Close
    Set cn = Nothing
    
End Sub

Sub FAC_Projets_Entête_Add_Record_To_DB(projetID As Long, _
                                        nomClient As String, _
                                        clientID As String, _
                                        dte As String, _
                                        hono As Double, _
                                        ByRef arr As Variant) 'Write a record to MASTER.xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Analyse:FAC_Projets_Entête_Add_Record_To_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_Entête$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE 1=0"
    
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conn, 2, 3
    
    Dim timeStamp As String
    Dim c As Long
    Dim l As Long
    rs.AddNew
        'Add fields to the recordset before updating it
        'RecordSet are ZERO base, and Enums are not, so the '-1' is mandatory !!!
        rs.Fields(fFacPEProjetID - 1).value = projetID
        rs.Fields(fFacPENomClient - 1).value = nomClient
        rs.Fields(fFacPEClientID - 1).value = clientID
        rs.Fields(fFacPEDate - 1).value = dte
        rs.Fields(fFacPEHonoTotal - 1).value = hono
        
        rs.Fields(fFacPEProf1 - 1).value = arr(1, 1)
        rs.Fields(fFacPEHres1 - 1).value = arr(1, 2)
        rs.Fields(fFacPETauxH1 - 1).value = arr(1, 3)
        rs.Fields(fFacPEHono1 - 1).value = arr(1, 4)
        
        If UBound(arr, 1) >= 2 Then
            rs.Fields(fFacPEProf2 - 1).value = arr(2, 1)
            rs.Fields(fFacPEHres2 - 1).value = arr(2, 2)
            rs.Fields(fFacPETauxH2 - 1).value = arr(2, 3)
            rs.Fields(fFacPEHono2 - 1).value = arr(2, 4)
        End If
        
        If UBound(arr, 1) >= 3 Then
            rs.Fields(fFacPEProf3 - 1).value = arr(3, 1)
            rs.Fields(fFacPEHres3 - 1).value = arr(3, 2)
            rs.Fields(fFacPETauxH3 - 1).value = arr(3, 3)
            rs.Fields(fFacPEHono3 - 1).value = arr(3, 4)
        End If
        
        If UBound(arr, 1) >= 4 Then
            rs.Fields(fFacPEProf4 - 1).value = arr(4, 1)
            rs.Fields(fFacPEHres4 - 1).value = arr(4, 2)
            rs.Fields(fFacPETauxH4 - 1).value = arr(4, 3)
            rs.Fields(fFacPEHono4 - 1).value = arr(4, 4)
        End If
        
        If UBound(arr, 1) >= 5 Then
            rs.Fields(fFacPEProf5 - 1).value = arr(5, 1)
            rs.Fields(fFacPEHres5 - 1).value = arr(5, 2)
            rs.Fields(fFacPETauxH5 - 1).value = arr(5, 3)
            rs.Fields(fFacPEHono5 - 1).value = arr(5, 4)
        End If
        
        rs.Fields(fFacPEestDetruite - 1).value = 0 'Faux
        timeStamp = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rs.Fields(fFacPETimeStamp - 1).value = timeStamp
    rs.Update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modTEC_Analyse:FAC_Projets_Entête_Add_Record_To_DB", "", startTime)

End Sub

Sub FAC_Projets_Entête_Add_Record_Locally(projetID As Long, nomClient As String, clientID As String, dte As String, hono As Double, ByRef arr As Variant) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Analyse:FAC_Projets_Entête_Add_Record_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in FAC_Projets_Détails?
    Dim lastUsedRow As Long, rn As Long
    lastUsedRow = wshFAC_Projets_Entête.Cells(wshFAC_Projets_Entête.Rows.count, "A").End(xlUp).row
    rn = lastUsedRow + 1
    
    Dim dateTEC As String, timeStamp As String
    wshFAC_Projets_Entête.Range("A" & rn).value = projetID
    wshFAC_Projets_Entête.Range("B" & rn).value = nomClient
    wshFAC_Projets_Entête.Range("C" & rn).value = clientID
    wshFAC_Projets_Entête.Range("D" & rn).value = dte
    wshFAC_Projets_Entête.Range("E" & rn).value = hono
    'Assign values from the array to the worksheet using .Cells
    Dim i As Long, j As Long
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            wshFAC_Projets_Entête.Cells(rn, 6 + (i - 1) * UBound(arr, 2) + j - 1).value = arr(i, j)
        Next j
    Next i
    wshFAC_Projets_Entête.Range("Z" & rn).value = "FAUX"
    timeStamp = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    wshFAC_Projets_Entête.Range("AA" & rn).value = timeStamp
    
    Application.ScreenUpdating = True

    Call Log_Record("modTEC_Analyse:FAC_Projets_Entête_Add_Record_Locally", "", startTime)

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
    strSQL = "UPDATE [" & sheetName & "] SET estDetruite = -1 WHERE [" & columnName & "] = '" & Replace(valueToFind, "'", "''") & "'"
    cn.Execute strSQL
    
    'Close the connection
    cn.Close
    Set cn = Nothing
    
End Sub

Sub Add_And_Modify_Checkbox(StartRow As Long, lastRow As Long)
    
    'Set your worksheet (adjust this to match your worksheet name)
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Define the range for the summary
    Dim summaryRange As Range
    Set summaryRange = ws.Range(ws.Cells(StartRow, 10), ws.Cells(lastRow, 13)) 'Columns J to M
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Add an ActiveX checkbox next to the summary in column O
    Dim checkBox As OLEObject
    With ws
        Set checkBox = .OLEObjects.Add(ClassType:="Forms.CheckBox.1", _
                    Left:=.Cells(lastRow, 14).Left + 5, _
                    Top:=.Cells(lastRow, 14).Top, Width:=80, Height:=16)
        
        'Modify checkbox properties
        With checkBox.Object
            .Caption = "On facture"
            .Font.size = 11  'Set font size
            .Font.Bold = True  'Set font bold
            .ForeColor = RGB(0, 0, 255)  'Set font color (Blue)
            .BackColor = RGB(200, 255, 200)  'Set background color (Light Green)
        End With
    End With
    
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set checkBox = Nothing
    Set summaryRange = Nothing
    Set ws = Nothing
    
End Sub

Sub Delete_CheckBox()

    'Set the worksheet
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Check if any CheckBox exists and then delete it/them
    Dim Sh As Shape
    For Each Sh In ws.Shapes
        If InStr(Sh.Name, "CheckBox") Then
            Sh.Delete
        End If
    Next Sh
    
    'Libérer la mémoire
    Set Sh = Nothing
    Set ws = Nothing
    
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
    ws.Rows(saveR & ":" & r).EntireRow.Hidden = True
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub Clear_Fees_Summary_And_CheckBox() 'RMV_15

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Analyse:Clear_Fees_Summary_And_CheckBox", "", 0)
    
    'Clean the Fees Summary Area
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    Application.EnableEvents = False
    ws.Range("J7:O9999").Clear
    Application.EnableEvents = True
    
    'Clear any leftover CheckBox
    Dim Sh As Shape
    For Each Sh In ws.Shapes
        If InStr(Sh.Name, "CheckBox") Then
            Sh.Delete
        End If
    Next Sh

    'Libérer la mémoire
    Set Sh = Nothing
    Set ws = Nothing
    
    Call Log_Record("modTEC_Analyse:Clear_Fees_Summary_And_CheckBox", "", startTime)
    
End Sub

Sub Get_CheckBox_Position(cb As OLEObject)

    'Set your worksheet (adjust this to match your worksheet name)
    Dim ws As Worksheet
    Set ws = wshTEC_Analyse
    
    'Reference your checkbox by name
    Dim checkBox As OLEObject
    Set checkBox = ws.OLEObjects(cb)
    
    'Get the cell that contains the top-left corner of the CheckBox
    Dim checkBoxCell As Range
    Set checkBoxCell = checkBox.TopLeftCell
    
    ' Display the address of the cell
    msgBox "The CheckBox is located at cell: " & checkBoxCell.Address
    
    'Libérer la mémoire
    Set checkBox = Nothing
    Set checkBoxCell = Nothing
    Set ws = Nothing

End Sub

Sub TEC_Analyse_Delete_CheckBox()

    'Assigner la feuille à ws
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Si CheckBox* existe, l'effacer
    Dim checkBox As OLEObject
    Dim i As Long
    For i = 1 To 5
        On Error Resume Next
        Set checkBox = ws.OLEObjects("CheckBox" & i)
        If Not checkBox Is Nothing Then
            checkBox.Delete
            
        End If
        On Error GoTo 0
    Next i
    
    'Libérer la mémoire
    Set checkBox = Nothing
    Set ws = Nothing
    
End Sub

Sub shp_TEC_Analyse_Back_To_TEC_Menu_Click()

    Call TEC_Analyse_Back_To_TEC_Menu

End Sub

Sub TEC_Analyse_Back_To_TEC_Menu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Analyse:TEC_Analyse_Back_To_TEC_Menu", "", 0)
    
    Call Clear_Fees_Summary_And_CheckBox
    
    Dim usedLastRow As Long
    usedLastRow = wshTEC_Analyse.Cells(wshTEC_Analyse.Rows.count, "C").End(xlUp).row
    Application.EnableEvents = False
    wshTEC_Analyse.Range("C6:O" & usedLastRow).Clear
    Application.EnableEvents = True
    
    wshTEC_Analyse.Visible = xlSheetVeryHidden
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select
    
    Call Log_Record("modTEC_Analyse:TEC_Analyse_Back_To_TEC_Menu", "", startTime)

End Sub

