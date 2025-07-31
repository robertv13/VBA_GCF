Attribute VB_Name = "modTEC_Analyse"
Option Explicit

Sub TEC_Sort_Group_And_Subtotal() '2024-08-24 @ 08:10

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:TEC_Sort_Group_And_Subtotal", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim wsDest As Worksheet: Set wsDest = wshTEC_Analyse
    
    'Remove existing subtotals in the destination worksheet
    wsDest.Cells.RemoveSubtotal
'    call modDev_utils.EnregistrerLogApplication("     modTEC_Analyse:TEC_Sort_Group_And_Subtotal - Les anciens SubTotal ont été effacés", -1)
    
    'Clear the worksheet from row 6 until the last row used
    Dim destLastUsedRow As Long
    destLastUsedRow = wsDest.Cells(wsDest.Rows.count, "B").End(xlUp).Row
    If destLastUsedRow < 6 Then destLastUsedRow = 6
    wsDest.Range("A6:I" & destLastUsedRow).Clear
    
    'Créer un dict pour tous les clients FACTURABLES
    Dim wsClientsMF As Worksheet: Set wsClientsMF = wsdBD_Clients
    Dim lastUsedRowClient As Long
    lastUsedRowClient = wsClientsMF.Cells(wsClientsMF.Rows.count, "B").End(xlUp).Row
    Dim dictClients As Dictionary
    Set dictClients = New Dictionary
    Dim clientData As Variant
    'Charger toutes les données des clients dans un tableau
    clientData = wsClientsMF.Range(wsClientsMF.Cells(2, fClntFMClientID), _
                               wsClientsMF.Cells(lastUsedRowClient, fClntFMClientNom)).Value

    'Parcourir le tableau pour ajouter les clients facturables au dictionnaire
    Dim i As Long
    For i = 1 To UBound(clientData, 1)
        If Fn_Is_Client_Facturable(clientData(i, 2)) = True Then
            dictClients.Add CStr(clientData(i, 2)), clientData(i, 1)
        End If
    Next i

    Dim lastUsedRow As Long, firstEmptyCol As Long
    
    'Set the source worksheet, lastUsedRow and lastUsedCol
    Dim wsSource As Worksheet: Set wsSource = wsdTEC_Local
    'Find the last row with data in the source worksheet
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).Row
    'Find the first empty column from the left in the source worksheet
    firstEmptyCol = 1
    Do Until IsEmpty(wsSource.Cells(2, firstEmptyCol))
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    Application.EnableEvents = False
    
    'Appel à AdvancedFilter # 2 dans TEC_Local
    Call Get_TEC_For_Client_AF(vbNullString, CLng(CDate(wsDest.Range("H3").Value)), "VRAI", "FAUX", "FAUX")
    
    Dim lastUsedResult As Long
    lastUsedResult = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "AQ").End(xlUp).Row
    
    'Charger les données sources dans un tableau (beaucoup plus rapide)
    Dim sourceData As Variant
    Dim rowCount As Long
'    Debug.Print "Il y a " & lastUsedResult & " rangées dans le tableau"
    sourceData = wsSource.Range("AQ3:AX" & lastUsedResult).Value
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
        wsDest.Range("A7:H" & r - 1 + 6).Value = outputData
        'Formats
        wsDest.Range("E7:F" & r - 1 + 6).HorizontalAlignment = xlCenter
        wsDest.Range("H7:H" & r - 1 + 6).NumberFormat = "#,##0.00"
    End If
    
    Application.EnableEvents = False
   
    'Find the last row in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).Row

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
    destLastUsedRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).Row
    Application.DisplayAlerts = False
    wsDest.Range("A6:H" & destLastUsedRow).Subtotal GroupBy:=3, Function:=xlSum, _
            TotalList:=Array(8), Replace:=True, PageBreaks:=False, SummaryBelowData:=False
    Application.DisplayAlerts = True
    wsDest.Range("A:B").EntireColumn.Hidden = True

    'Group the data to show subtotals in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).Row
    wsDest.Outline.ShowLevels RowLevels:=2
    
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
        If wsDest.Range("A" & r).Value = vbNullString Then
            With wsDest.Range("C" & r).Interior
'                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
'                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
'                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.399975585192419
'                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
'                .PatternTintAndShade = 0
            End With
            With wsDest.Range("C" & r).Font
                .ThemeColor = xlThemeColorLight1
'                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            With wsDest.Range("C" & r)
'                If InStr(.Value, "Total ") = 1 Then
'                    .Value = Mid$(.Value, 7)
'                End If
                If .Value = "Total général" Then
                    .Value = "G r a n d   T o t a l"
                End If
            End With
            'Mettre de l'emphase sur les cellules d'heures, si le montant du projet <> 0,00 $
            If wsDest.Range("D" & r).Value = 0 Then
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
    Call AppliquerFormatConditionnelZebreeColonneH(rngTotals, destLastUsedRow)
    
    'Bring in all the invoice requests
    Call ChargerDemandesDeFactureExistantes(destLastUsedRow)
    
    'Clean up the summary area of the worksheet
    Call EffacerPlageSommaire(wsDest)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set dictClients = Nothing
    Set rngTotals = Nothing
    Set wsClientsMF = Nothing
    Set wsDest = Nothing
    Set wsSource = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:TEC_Sort_Group_And_Subtotal", vbNullString, startTime)

End Sub

Sub EffacerPlageSommaire(ws As Worksheet)

    Application.EnableEvents = False
    
    'Clean up the summary area (columns K to Q)
    ws.Range("J:P").Clear
    'Erase any checkbox left over
    Call EffacerCheckBox
    
    Application.EnableEvents = True

End Sub

Sub AppliquerFormatConditionnelZebreeColonneH(rng As Range, lastUsedRow As Long)

    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Loop each cell in column C to find Totals row
    Dim totalRange As Range, cell As Range
    For Each cell In rng
        If InStr(1, cell.Value, "Total ", vbTextCompare) > 0 Then
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
            With .item(.Count)
                .Interior.Color = RGB(255, 0, 0) 'Red color
            End With
    
            'Rule for values > 25
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="25"
            With .item(.Count)
                .Interior.Color = RGB(255, 165, 0) 'Orange color
            End With
    
            'Rule for values > 10
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="10"
            With .item(.Count)
                .Interior.Color = RGB(255, 255, 0) 'Yellow color
            End With
    
            'Rule for values > 5
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="5"
            With .item(.Count)
                .Interior.Color = RGB(144, 238, 144) 'Light green color
            End With
        End With
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set totalRange = Nothing
    Set ws = Nothing
            
End Sub

Sub ConstruireSommaireHeures(rowSelected As Long)

    If rowSelected < 8 Then Exit Sub
    
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Determine the last row used
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    'Clear the Hours Summary area
    Call EffacerPlageSommaire(ws)
    
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
    Dim i As Long, saveR As Long
    rowSelected = rowSelected + 1 'Summary starts on the next line (first line of expanded lines)
    saveR = rowSelected
    i = rowSelected
    Do Until ActiveSheet.Cells(i, 5) = vbNullString
        If ActiveSheet.Cells(i, 6).Value <> vbNullString Then
            If dictHours.Exists(ActiveSheet.Cells(i, 6).Value) Then
                dictHours(ActiveSheet.Cells(i, 6).Value) = dictHours(ActiveSheet.Cells(i, 6).Value) + ActiveSheet.Cells(i, 8).Value
            Else
                dictHours.Add ActiveSheet.Cells(i, 6).Value, ActiveSheet.Cells(i, 8).Value
            End If
            ActiveSheet.Cells(i, 8).Font.Color = RGB(166, 166, 166) 'RMV_15
        End If
        i = i + 1
    Loop

    Dim prof As Variant
    Dim profID As Long
    Dim tauxHoraire As Currency
    
    Application.EnableEvents = False
    
    ws.Range("O" & rowSelected).Value = 0 'Reset the total WIP value
    For Each prof In Fn_Sort_Dictionary_By_Value(dictHours, True) ' Sort dictionary by hours in descending order
        ActiveSheet.Cells(rowSelected, 10).Value = prof
        Dim strProf As String
        strProf = prof
        profID = Fn_GetID_From_Initials(strProf)
        ActiveSheet.Cells(rowSelected, "K").HorizontalAlignment = xlRight
        ActiveSheet.Cells(rowSelected, "K").NumberFormat = "#,##0.00"
        ActiveSheet.Cells(rowSelected, "K").Value = dictHours(prof)
        tauxHoraire = Fn_Get_Hourly_Rate(profID, ws.Range("H3").Value)
        ActiveSheet.Cells(rowSelected, "L").Value = tauxHoraire
        ActiveSheet.Cells(rowSelected, "M").NumberFormat = "#,##0.00 $"
        ActiveSheet.Cells(rowSelected, "M").FormulaR1C1 = "=RC[-2]*RC[-1]"
        ActiveSheet.Cells(rowSelected, "M").HorizontalAlignment = xlRight
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
    With ActiveSheet.Cells(rTotal, "K")
        .HorizontalAlignment = xlRight
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
'        .Value = Format$(t, "#,##0.00")
        .Font.Bold = True
    End With
    
    'Fees Total
    With ActiveSheet.Cells(rowSelected, "M")
        .HorizontalAlignment = xlRight
'        .Value = Format$(tdollars, "#,##0.00$")
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
        .Font.Bold = True
    End With
    
    With ActiveSheet.Range("J" & saveR & ":M" & rowSelected).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    With ActiveSheet.Range("K" & rowSelected & ", M" & rowSelected)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With

    'Save the TOTAL WIP value
    With ws.Range("N" & saveR)
        .Value = "Valeur TEC:"
        .Font.Italic = True
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("O" & saveR)
        .NumberFormat = "#,##0.00 $"
        .Value = ws.Range("M" & rowSelected).Value
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
    
    Call AjouterCaseACocherOnFacture(saveR, rowSelected)
    
    Application.EnableEvents = True

    'Libérer la mémoire
    Set dictHours = Nothing
    Set prof = Nothing
    Set rngSort = Nothing
    Set ws = Nothing
    
End Sub
    
Sub ChargerDemandesDeFactureExistantes(activeLastUsedRow As Long)

    Dim wsSource As Worksheet: Set wsSource = wsdFAC_Projets_Entete
    Dim sourceLastUsedRow As Long
    sourceLastUsedRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    
    Dim wsActive As Worksheet: Set wsActive = wshTEC_Analyse
    Dim rngTotal As Range: Set rngTotal = wsActive.Range("C1:C" & activeLastUsedRow)
    
    'Analyze all Invoice Requests (one row at the time)
    
    Dim clientName As String
    Dim clientID As String
    Dim honoTotal As Double
    Dim result As Variant
    Dim i As Long, r As Long
    For i = 2 To sourceLastUsedRow
        If wsSource.Cells(i, 26).Value <> "True" And wsSource.Cells(i, 26).Value <> -1 Then
            clientName = wsSource.Cells(i, 2).Value
            clientID = wsSource.Cells(i, 3).Value
            honoTotal = wsSource.Cells(i, 5).Value
            'Using XLOOKUP to find the result directly
            result = Application.WorksheetFunction.XLookup("Total " & clientName, _
                                                           rngTotal, _
                                                           rngTotal, _
                                                           "Not Found", _
                                                           0, _
                                                           1)
            
            If result <> "Not Found" Then
                r = Application.WorksheetFunction.Match(result, rngTotal, 0)
                wsActive.Cells(r, 4).Value = honoTotal
                wsActive.Cells(r, 4).NumberFormat = "#,##0.00 $"
            End If
        End If
    Next i

    'Libérer la mémoire
    Set rngTotal = Nothing
    Set wsActive = Nothing
    Set wsSource = Nothing
    
End Sub

Sub FAC_Projets_Details_Add_Record_To_DB(clientID As String, fr As Long, lr As Long, ByRef projetID As Long) 'Write a record to MASTER.xlsx file
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:FAC_Projets_Details_Add_Record_To_DB", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Projets_Details$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")
    
    'First SQL - SQL query to find the maximum value in the first column
    Dim strSQL As String
    strSQL = "SELECT MAX(ProjetID) AS MaxValue FROM [" & destinationTab & "]"
    recSet.Open strSQL, conn

    'Get the maximum value
    Dim MaxValue As Long
    If IsNull(recSet.Fields("MaxValue").Value) Then
        'Handle empty table (assign a default value, e.g., 1)
        projetID = 1
    Else
        projetID = recSet.Fields("MaxValue").Value + 1
    End If
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset (no longer needed)
    recSet.Close
    
    'Second SQL - SQL query to add the new records
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE 1=0"
    recSet.Open strSQL, conn, 2, 3
    
    'Read all line from TEC_Analyse
    Dim dateTEC As String
    Dim l As Long
    For l = fr To lr
        recSet.AddNew
            'Add fields to the recordset before updating it
            'RecordSet are ZERO base, and Enums are not, so the '-1' is mandatory !!!
            recSet.Fields(fFacPDProjetID - 1).Value = projetID
            recSet.Fields(fFacPDNomClient - 1).Value = wshTEC_Analyse.Range("C" & l).Value
            recSet.Fields(fFacPDClientID - 1).Value = CStr(clientID)
            recSet.Fields(fFacPDTECID - 1).Value = wshTEC_Analyse.Range("A" & l).Value
            recSet.Fields(fFacPDProfID - 1).Value = wshTEC_Analyse.Range("B" & l).Value
            dateTEC = Format$(wshTEC_Analyse.Range("E" & l).Value, "yyyy-mm-dd")
            recSet.Fields(fFacPDDate - 1).Value = dateTEC
            recSet.Fields(fFacPDProf - 1).Value = wshTEC_Analyse.Range("F" & l).Value
            recSet.Fields(fFacPDestDetruite - 1) = 0 'Faux
            recSet.Fields(fFacPDHeures - 1).Value = CDbl(wshTEC_Analyse.Range("H" & l).Value)
            recSet.Fields(fFacPDTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        recSet.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:FAC_Projets_Details_Add_Record_To_DB", vbNullString, startTime)

End Sub

Sub FAC_Projets_Details_Add_Record_Locally(clientID As String, fr As Long, lr As Long, projetID As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:FAC_Projets_Details_Add_Record_Locally", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in FAC_Projets_Details?
    Dim lastUsedRow As Long, rn As Long
    lastUsedRow = wsdFAC_Projets_Details.Cells(wsdFAC_Projets_Details.Rows.count, "A").End(xlUp).Row
    If wsdFAC_Projets_Details.Cells(2, 1).Value = vbNullString Then
        rn = lastUsedRow
    Else
        rn = lastUsedRow + 1
    End If
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim dateTEC As String
    Dim i As Long
    For i = fr To lr
        wsdFAC_Projets_Details.Range("A" & rn).Value = projetID
        wsdFAC_Projets_Details.Range("B" & rn).Value = wshTEC_Analyse.Range("C" & i).Value
        wsdFAC_Projets_Details.Range("C" & rn).Value = clientID
        wsdFAC_Projets_Details.Range("D" & rn).Value = wshTEC_Analyse.Range("A" & i).Value
        wsdFAC_Projets_Details.Range("E" & rn).Value = wshTEC_Analyse.Range("B" & i).Value
        dateTEC = Format$(wshTEC_Analyse.Range("E" & i).Value, "yyyy-mm-dd")
        wsdFAC_Projets_Details.Range("F" & rn).Value = dateTEC
        wsdFAC_Projets_Details.Range("G" & rn).Value = wshTEC_Analyse.Range("F" & i).Value
        wsdFAC_Projets_Details.Range("H" & rn).Value = wshTEC_Analyse.Range("H" & i).Value
        wsdFAC_Projets_Details.Range("I" & rn).Value = "FAUX"
        wsdFAC_Projets_Details.Range("J" & rn).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        rn = rn + 1
    Next i
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:FAC_Projets_Details_Add_Record_Locally", vbNullString, startTime)

End Sub

Sub DetruireDetailSiEnteteEstDetruite(filePath As String, _
                                                    sheetName As String, _
                                                    columnName As String, _
                                                    valueToFind As Variant) '2024-07-19 @ 15:31
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:DetruireDetailSiEnteteEstDetruite", vbNullString, 0)
    
    'Create a new ADODB connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    'Open the connection to the closed workbook
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";" & _
              "Extended Properties=""Excel 12.0;HDR=Yes"";"
    
    'Update the rows to mark as deleted (soft delete)
    Dim strSQL As String
    strSQL = "UPDATE [" & sheetName & "] SET estDetruite = -1 WHERE [" & columnName & "] = '" & Replace(valueToFind, "'", "''") & "'"
    conn.Execute strSQL
    
    'Close the connection
    conn.Close
    Set conn = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:DetruireDetailSiEnteteEstDetruite", vbNullString, startTime)

End Sub

Sub FAC_Projets_Entete_Add_Record_To_DB(projetID As Long, _
                                        nomClient As String, _
                                        clientID As String, _
                                        dte As String, _
                                        hono As Double, _
                                        ByRef arr As Variant) 'Write a record to MASTER.xlsx file
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:FAC_Projets_Entete_Add_Record_To_DB", vbNullString, 0)
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Projets_Entete$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")
    
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE 1=0"
    
    recSet.Open strSQL, conn, 2, 3
    
    Dim c As Long
    Dim l As Long
    recSet.AddNew
        Debug.Print "recSet.state = " & recSet.state
        Debug.Print "recSet.EOF = " & recSet.EOF
        Debug.Print "recSet.BOF = " & recSet.BOF
        Debug.Print "recSet.RecordCount = " & recSet.RecordCount
        Debug.Print "recSet.Fields.Count = " & recSet.Fields.count
        Debug.Print "Field 2 - Name: " & recSet.Fields(2).Name
        Debug.Print "Field 2 - Type: " & recSet.Fields(2).Type
        Debug.Print "Field 2 - Attributes: " & recSet.Fields(2).Attributes
        'Add fields to the recordset before updating it
        'RecordSet are ZERO base, and Enums are not, so the '-1' is mandatory !!!
        recSet.Fields(fFacPEProjetID - 1).Value = projetID
        recSet.Fields(fFacPENomClient - 1).Value = nomClient
        recSet.Fields(fFacPEClientID - 1).Value = CStr(clientID)
        recSet.Fields(fFacPEDate - 1).Value = dte
        recSet.Fields(fFacPEHonoTotal - 1).Value = hono
        
        recSet.Fields(fFacPEProf1 - 1).Value = arr(1, 1)
        recSet.Fields(fFacPEHres1 - 1).Value = arr(1, 2)
        recSet.Fields(fFacPETauxH1 - 1).Value = arr(1, 3)
        recSet.Fields(fFacPEHono1 - 1).Value = arr(1, 4)
        
        If UBound(arr, 1) >= 2 Then
            recSet.Fields(fFacPEProf2 - 1).Value = arr(2, 1)
            recSet.Fields(fFacPEHres2 - 1).Value = arr(2, 2)
            recSet.Fields(fFacPETauxH2 - 1).Value = arr(2, 3)
            recSet.Fields(fFacPEHono2 - 1).Value = arr(2, 4)
        End If
        
        If UBound(arr, 1) >= 3 Then
            recSet.Fields(fFacPEProf3 - 1).Value = arr(3, 1)
            recSet.Fields(fFacPEHres3 - 1).Value = arr(3, 2)
            recSet.Fields(fFacPETauxH3 - 1).Value = arr(3, 3)
            recSet.Fields(fFacPEHono3 - 1).Value = arr(3, 4)
        End If
        
        If UBound(arr, 1) >= 4 Then
            recSet.Fields(fFacPEProf4 - 1).Value = arr(4, 1)
            recSet.Fields(fFacPEHres4 - 1).Value = arr(4, 2)
            recSet.Fields(fFacPETauxH4 - 1).Value = arr(4, 3)
            recSet.Fields(fFacPEHono4 - 1).Value = arr(4, 4)
        End If
        
        If UBound(arr, 1) >= 5 Then
            recSet.Fields(fFacPEProf5 - 1).Value = arr(5, 1)
            recSet.Fields(fFacPEHres5 - 1).Value = arr(5, 2)
            recSet.Fields(fFacPETauxH5 - 1).Value = arr(5, 3)
            recSet.Fields(fFacPEHono5 - 1).Value = arr(5, 4)
        End If
        
        recSet.Fields(fFacPEestDetruite - 1).Value = 0 'Faux
        recSet.Fields(fFacPETimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    recSet.Update
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:FAC_Projets_Entete_Add_Record_To_DB", vbNullString, startTime)

End Sub

Sub FAC_Projets_Entete_Add_Record_Locally(projetID As Long, nomClient As String, clientID As String, dte As String, hono As Double, ByRef arr As Variant) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:FAC_Projets_Entete_Add_Record_Locally", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in FAC_Projets_Details?
    Dim lastUsedRow As Long, rn As Long
    lastUsedRow = wsdFAC_Projets_Entete.Cells(wsdFAC_Projets_Entete.Rows.count, "A").End(xlUp).Row
    If wsdFAC_Projets_Entete.Cells(2, 1).Value = vbNullString Then
        rn = lastUsedRow
    Else
        rn = lastUsedRow + 1
    End If
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim dateTEC As String
    wsdFAC_Projets_Entete.Range("A" & rn).Value = projetID
    wsdFAC_Projets_Entete.Range("B" & rn).Value = nomClient
    wsdFAC_Projets_Entete.Range("C" & rn).Value = clientID
    wsdFAC_Projets_Entete.Range("D" & rn).Value = dte
    wsdFAC_Projets_Entete.Range("E" & rn).Value = hono
    'Assign values from the array to the worksheet using .Cells
    Dim i As Long, j As Long
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            wsdFAC_Projets_Entete.Cells(rn, 6 + (i - 1) * UBound(arr, 2) + j - 1).Value = arr(i, j)
        Next j
    Next i
    wsdFAC_Projets_Entete.Range("Z" & rn).Value = "FAUX"
    wsdFAC_Projets_Entete.Range("AA" & rn).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:FAC_Projets_Entete_Add_Record_Locally", vbNullString, startTime)

End Sub

Sub DetruireEnteteSiEnteteEstDetruite(filePath As String, _
                                      sheetName As String, _
                                      columnName As String, _
                                      valueToFind As Variant) '2024-07-19 @ 15:31
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:DetruireEnteteSiEnteteEstDetruite", vbNullString, 0)
    
    'Create a new ADODB connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";" & _
              "Extended Properties=""Excel 12.0;HDR=Yes"";"
    
    'Update the rows to mark as deleted (soft delete)
    Dim strSQL As String
    strSQL = "UPDATE [" & sheetName & "] SET estDetruite = -1 WHERE [" & columnName & "] = '" & Replace(valueToFind, "'", "''") & "'"
    conn.Execute strSQL
    
    'Close the connection
    conn.Close
    Set conn = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:DetruireEnteteSiEnteteEstDetruite", vbNullString, startTime)

End Sub

Sub AjouterCaseACocherOnFacture(StartRow As Long, lastRow As Long)
    
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

Sub EffacerCheckBox()

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
    Do While wshTEC_Analyse.Range("A" & r).Value <> vbNullString
        r = r + 1
    Loop

    r = r - 1
    ws.Rows(saveR & ":" & r).EntireRow.Hidden = True
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub EffacerSectionHonorairesEtCheckBox() 'RMV_15

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:EffacerSectionHonorairesEtCheckBox", vbNullString, 0)
    
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
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:EffacerSectionHonorairesEtCheckBox", vbNullString, startTime)
    
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

Sub NettoyerProjetsDetruits(loDetails As ListObject, loEntete As ListObject) '2025-07-11 @ 01:50

    Dim i As Long
    Dim colDetruite_Entete As Long: colDetruite_Entete = 26
    Dim colDetruite_Detail As Long: colDetruite_Detail = 9

    'D'abord les DÉTAILS — important de commencer par les enfants
    For i = loDetails.ListRows.count To 1 Step -1
        With loDetails.ListRows(i).Range.Cells(1, colDetruite_Detail)
            If LCase(Trim(.Value)) = "vrai" Or .Value = True Or .Value = -1 Then
                loDetails.ListRows(i).Delete
            End If
        End With
    Next i

    'Ensuite les ENTÊTES
    For i = loEntete.ListRows.count To 1 Step -1
        With loEntete.ListRows(i).Range.Cells(1, colDetruite_Entete)
            If LCase(Trim(.Value)) = "vrai" Or .Value = True Or .Value = -1 Then
                loEntete.ListRows(i).Delete
            End If
        End With
    Next i

End Sub

Sub shp_TEC_Analyse_Back_To_TEC_Menu_Click()

    Call TEC_Analyse_Back_To_TEC_Menu

End Sub

Sub TEC_Analyse_Back_To_TEC_Menu()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:TEC_Analyse_Back_To_TEC_Menu", vbNullString, 0)
    
    Dim loDetails As ListObject
    Set loDetails = wsdFAC_Projets_Details.ListObjects("l_tbl_FAC_Projets_Details")
    Dim loEntete As ListObject
    Set loEntete = wsdFAC_Projets_Entete.ListObjects("l_tbl_FAC_Projets_Entete")
    
    Call NettoyerProjetsDetruits(loDetails, loEntete)
    
    Call EffacerSectionHonorairesEtCheckBox
    
    Dim usedLastRow As Long
    usedLastRow = wshTEC_Analyse.Cells(wshTEC_Analyse.Rows.count, "C").End(xlUp).Row
    Application.EnableEvents = False
    wshTEC_Analyse.Range("C7:O" & usedLastRow).Clear
    Application.EnableEvents = True
    
    wshTEC_Analyse.Visible = xlSheetVeryHidden
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Analyse:TEC_Analyse_Back_To_TEC_Menu", vbNullString, startTime)

End Sub


