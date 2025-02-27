Attribute VB_Name = "modzDataConversion"
Option Explicit

'Importation des clients à partir de ... \DataConversion\Clients.xlsx
Sub CopieClientsEntreClasseursFermés() '2024-08-03 @ 09:40

    Stop 'One shot deal !!!
    
    Dim sourceRange As Range
    
    'Définir les chemins d'accès des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALITÉ\DataConversion\Clients.xlsx"
    Dim destinationFilePath As String
    destinationFilePath = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & "GCF_BD_Entrée.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("Feuil1")
    
    'Détermine la dernière rangée utilisée dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.Rows.count, 1).End(xlUp).row
    Dim lastUsedCol As Long
    lastUsedCol = sourceSheet.Cells(1, sourceSheet.Columns.count).End(xlToLeft).Column
    
    'Define the range to copy
    Set sourceRange = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastUsedRow, lastUsedCol))
    
    ' Open the destination workbook
    Dim destinationWorkbook As Workbook: Set destinationWorkbook = Workbooks.Open(destinationFilePath)
    Dim destinationSheet As Worksheet: Set destinationSheet = destinationWorkbook.Worksheets("Clients")
    
    'Clear existing data in the destination sheet
    destinationSheet.Cells.ClearContents
    
    'Copy the data from the source to the destination
    sourceRange.Copy Destination:=destinationSheet.Range("A1")
    
    'Save and close the destination workbook
    destinationWorkbook.Save
    destinationWorkbook.Close
    
    'Close the source workbook without saving
    sourceWorkbook.Close SaveChanges:=False
    
    'Libérer la mémoire
    Set sourceSheet = Nothing
    Set sourceRange = Nothing
    Set sourceWorkbook = Nothing
    Set destinationSheet = Nothing
    Set destinationWorkbook = Nothing
    
    MsgBox "Les données ont été copiées avec succès dans le fichier destination."
    
End Sub

'Ajustements à la feuille DB_Clients (*) ---> [*]
Sub AjusteNomClient()

    'Declare and open the closed workbook
    Dim wb As Workbook: Set wb = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx")

    'Define the worksheet you want to work with
    Dim ws As Worksheet: Set ws = wb.Worksheets("Clients")
    
    'Find the last used row with data in column A
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    'Loop through each row starting from row 2 (headers are 1 row)
    Dim client As String, clientID As String, contactFacturation As String
    Dim posOpenParenthesis As Integer, posCloseParenthesis As Integer
    Dim numberOpenParenthesis As Integer, numberCloseParenthesis As Integer
    Dim i As Long
    For i = 2 To lastUsedRow
        'Load data into variables
        client = ws.Cells(i, fClntFMClientNom).value
        clientID = ws.Cells(i, fClntFMClientID).value
        contactFacturation = ws.Cells(i, fClntFMContactFacturation).value
        
        'Process the data and make adjustments if necessary
        posOpenParenthesis = InStr(client, "(")
        posCloseParenthesis = InStr(client, ")")
        numberOpenParenthesis = Fn_Count_Char_Occurrences(client, "(")
        numberCloseParenthesis = Fn_Count_Char_Occurrences(client, ")")
        
        If numberOpenParenthesis = 1 And numberCloseParenthesis = 1 Then
            If posCloseParenthesis > posOpenParenthesis + 5 Then
                client = Replace(client, "(", "[")
                client = Replace(client, ")", "]")
                ws.Cells(i, 1).value = client
                Debug.Print "#064 - " & i & " - " & client
            End If
        End If
        
    Next i
    
    wb.Save
    
    'Libérer la mémoire
    Set wb = Nothing
    Set ws = Nothing
    
    MsgBox "Le traitement est complété sur " & i - 1 & " lignes"
    
End Sub

'Ajustements à la feuille DB_Clients (Ajout du contactdans le nom du client)
Sub AjouteContactDansNomClient()

    'Declare and open the closed workbook
    Dim wb As Workbook: Set wb = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx")

    'Define the worksheet you want to work with
    Dim ws As Worksheet: Set ws = wb.Worksheets("Clients")
    
    'Find the last used row with data in column A
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    'Loop through each row starting from row 2 (headers are 1 row)
    Dim client As String, clientID As String, contactFacturation As String
    Dim posOpenSquareBracket As Integer, posCloseSquareBracket As Integer
    Dim numberOpenSquareBracket As Integer, numberCloseSquareBracket As Integer
    Dim i As Long
    For i = 2 To lastUsedRow
        'Load data into variables
        client = ws.Cells(i, fClntFMClientNom).value
        clientID = ws.Cells(i, fClntFMClientID).value
        contactFacturation = Trim(ws.Cells(i, fClntFMContactFacturation).value)
        
        'Process the data and make adjustments if necessary
        posOpenSquareBracket = InStr(client, "[")
        posCloseSquareBracket = InStr(client, "]")
        
        If numberOpenSquareBracket = 0 And numberCloseSquareBracket = 0 Then
            If contactFacturation <> "" And InStr(client, contactFacturation) = 0 Then
                client = Trim(client) & " [" & contactFacturation & "]"
                ws.Cells(i, 1).value = client
                Debug.Print "#065 - " & i & " - " & client
            End If
        End If
        
    Next i
    
    wb.Save
    
    'Libérer la mémoire
    Set wb = Nothing
    Set ws = Nothing
    
    MsgBox "Le traitement est complété sur " & i - 1 & " lignes"
    
End Sub

Sub ImporterDonnéesDeClasseursFermés_TEC() '2024-08-14 @ 06:43 & 2024-08-03 @ 16:15

    Stop 'One shot deal !!!
    
    Call Client_List_Import_All
    
    'Define the path to the closed workbook
    Dim strFilePath As String
    strFilePath = "C:\VBA\GC_FISCALITÉ\DataConversion\TEC_20240814.xlsx"
    Dim strSheetName As String
    strSheetName = "TEC$"
    Dim strRange As String
    strRange = "A1:F110" 'Adjust the range as needed
    
    'Connection string for Excel
    Dim strConnection As String
    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & strFilePath & ";" & _
                    "Extended Properties=""Excel 12.0 Xml;HDR=Yes"";"
    
    'Create a new ADO connection and recordset
    Dim cnn As Object: Set cnn = CreateObject("ADODB.Connection")
    Dim rst As Object: Set rst = CreateObject("ADODB.Recordset")
    
    'Open the connection
    cnn.Open strConnection
    
    'Open the recordset
    rst.Open "SELECT * FROM [" & strSheetName & strRange & "]", cnn, 3, 1, 1
    
    'Define the destination worksheet
    Dim wsDest As Worksheet
    Set wsDest = wshTEC_Local
    
    'Get the last row in the destination sheet
    Dim lastUsedRow As Long, rowNum As Long
    lastUsedRow = wsDest.Cells(wsDest.Rows.count, "A").End(xlUp).row
    rowNum = lastUsedRow
    
    'Loop through the recordset and write data to the destination sheet
    Dim prof As String
    Dim client As String
    Dim clientCode As String
    Dim clientCodeFromDB As String
    Dim errorMesg As String
    Dim tecID As Long: tecID = 342
    Dim totHres As Double
    Do Until rst.EOF
        rowNum = rowNum + 1
        prof = Trim(rst.Fields(0).value)
        clientCode = Trim(rst.Fields(2).value)
'        clientCode = Left(client, 10)
'            clientCode = Left(clientCode, InStr(clientCode, " -") - 1)
        client = Trim(rst.Fields(3).value)
'        client = Mid(client, InStr(client, " - ") + 3, Len(client))
        totHres = totHres + CDbl(rst.Fields(5).value)
        
        'Is this a Valid Client ?
        Dim myInfo() As Variant
        Dim rng As Range: Set rng = wshBD_Clients.Range("dnrClients_Names_Only")
        myInfo = Fn_Find_Data_In_A_Range(rng, 2, clientCode, 1)
        If myInfo(1) = "" Then
            If InStr(errorMesg, client) = 0 Then
                errorMesg = errorMesg & clientCode & " - " & client & vbNewLine
            End If
        Else
            client = myInfo(3)
        End If
        
        tecID = tecID + 1
        wsDest.Range("A" & rowNum).value = tecID
        wsDest.Range("B" & rowNum).value = ObtenirProfIDAvecInitiales(prof)
        wsDest.Range("C" & rowNum).value = prof
        wsDest.Range("D" & rowNum).value = rst.Fields(1).value
        wsDest.Range("E" & rowNum).value = clientCode
        wsDest.Range("F" & rowNum).value = client
        wsDest.Range("G" & rowNum).value = rst.Fields(4).value
        wsDest.Range("H" & rowNum).value = rst.Fields(5).value
        wsDest.Range("I" & rowNum).value = ""
        wsDest.Range("J" & rowNum).value = "VRAI"
        wsDest.Range("K" & rowNum).value = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
        wsDest.Range("L" & rowNum).value = "FAUX"
        wsDest.Range("M" & rowNum).value = ""
        wsDest.Range("N" & rowNum).value = "FAUX"
        wsDest.Range("O" & rowNum).value = ThisWorkbook.Name
        wsDest.Range("P" & rowNum).value = ""
        
        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les TEC ont été importés, pour un total de " & totHres & " heures"
    End If
    
    'Libérer la mémoire
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Set rng = Nothing
    Set wsDest = Nothing
    
End Sub

'Only valid for this conversion process
Function ObtenirProfIDAvecInitiales(p As String) As Long

    Stop 'One shot deal
    
    Select Case p
        Case "GC"
            ObtenirProfIDAvecInitiales = 1
        Case "VG"
            ObtenirProfIDAvecInitiales = 2
        Case "AR"
            ObtenirProfIDAvecInitiales = 3
        Case "ML"
            ObtenirProfIDAvecInitiales = 4
        Case Else
            ObtenirProfIDAvecInitiales = 0
    End Select

End Function

'Importation des fournisseurs à partir de ... \DataConversion\Fournisseurs.xlsx
Sub ImporterDonnéesDeClasseursFermésFournisseurs() '2024-08-03 @ 18:10

    Stop 'One shot deal
    
    'Définir les chemins d'accès des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALITÉ\DataConversion\Fournisseurs.xlsx"
    Dim destinationFilePath As String
    destinationFilePath = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & "GCF_BD_Entrée.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("Fournisseurs")
    
    'Détermine la dernière rangée utilisée dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.Rows.count, 1).End(xlUp).row
    Dim lastUsedCol As Long
    lastUsedCol = sourceSheet.Cells(1, sourceSheet.Columns.count).End(xlToLeft).Column
    
    'Define the range to copy
    Dim sourceRange As Range
    Set sourceRange = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastUsedRow, lastUsedCol))
    
    ' Open the destination workbook
    Dim destinationWorkbook As Workbook: Set destinationWorkbook = Workbooks.Open(destinationFilePath)
    Dim destinationSheet As Worksheet: Set destinationSheet = destinationWorkbook.Worksheets("Fournisseurs")
    
    'Clear existing data in the destination sheet
    destinationSheet.Cells.ClearContents
    
    'Copy the data from the source to the destination
    sourceRange.Copy Destination:=destinationSheet.Range("A1")
    
    'Save and close the destination workbook
    destinationWorkbook.Save
    destinationWorkbook.Close
    
    'Close the source workbook without saving
    sourceWorkbook.Close SaveChanges:=False
    
    'Libérer la mémoire
    Set sourceSheet = Nothing
    Set sourceRange = Nothing
    Set sourceWorkbook = Nothing
    Set destinationSheet = Nothing
    Set destinationWorkbook = Nothing
    
    MsgBox "Les données (fournisseurs) ont été copiées avec succès dans" & vbNewLine & _
            vbNewLine & "le fichier destination."
    
End Sub

Sub ImporterDonnéesDeClasseursFermés_GL_BV() '2024-08-03 @ 18:20

    Stop 'One shot deal
    
    'Define the path to the closed workbook
    Dim strFilePath As String
    strFilePath = "C:\VBA\GC_FISCALITÉ\DataConversion\GL_BV.xlsx"
    Dim strSheetName As String
    strSheetName = "BV$"
    Dim strRange As String
    strRange = "A1:B20" 'Adjust the range as needed
    
    'Connection string for Excel
    Dim strConnection As String
    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & strFilePath & ";" & _
                    "Extended Properties=""Excel 12.0 Xml;HDR=Yes"";"
    
    'Create a new ADO connection and recordset
    Dim cnn As Object: Set cnn = CreateObject("ADODB.Connection")
    Dim rst As Object: Set rst = CreateObject("ADODB.Recordset")
    
    'Open the connection
    cnn.Open strConnection
    
    'Open the recordset
    rst.Open "SELECT * FROM [" & strSheetName & strRange & "]", cnn, 3, 1, 1
    
    'Define the destination worksheet
    Dim wsDest As Worksheet: Set wsDest = wshGL_Trans
    
    'Get the last row in the destination sheet
    Dim lastUsedRow As Long
    lastUsedRow = wsDest.Cells(wsDest.Rows.count, "A").End(xlUp).row
    Dim rowNum As Long
    rowNum = lastUsedRow
    
    'Loop through the recordset and write data to the destination sheet
    Dim descriptionGL As String
    Dim codeGL As String
    Dim errorMesg As String
    Dim No_Entrée As Long
    No_Entrée = 1
    Dim amount As Double
    Dim totalDT As Double, totalCT As Double
    
    Do Until rst.EOF
        rowNum = rowNum + 1
        descriptionGL = rst.Fields(0).value
        codeGL = Fn_GetGL_Code_From_GL_Description(descriptionGL)
        amount = rst.Fields(1).value
        If amount > 0 Then
            totalDT = totalDT + amount
        Else
            totalCT = totalCT - amount
        End If
        
        wsDest.Range("A" & rowNum).value = No_Entrée
        wsDest.Range("B" & rowNum).value = "07/31/2024"
        wsDest.Range("C" & rowNum).value = "Solde de fermeture (conversion)"
        wsDest.Range("D" & rowNum).value = "Conv."
        wsDest.Range("E" & rowNum).value = codeGL
        wsDest.Range("F" & rowNum).value = descriptionGL
        If amount >= 0 Then
            wsDest.Range("G" & rowNum).value = amount
        Else
            wsDest.Range("H" & rowNum).value = -amount
        End If
        wsDest.Range("I" & rowNum).value = ""
        wsDest.Range("J" & rowNum).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        
        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les soldes ont été importés, pour un total débit de " & totalDT & _
                vbNewLine & vbNewLine & "un total crédit de " & totalCT
    End If
    
    'Libérer la mémoire
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Set wsDest = Nothing
    
End Sub

Sub ImporterDonnéesDeClasseursFermés_CAR() '2024-08-04 @ 07:31

    Stop 'One shot deal
    
    Call Client_List_Import_All
    
    Dim strConnection As String
    Dim wsDest As Worksheet
    Dim i As Long, j As Long
    Dim lastUsedRow As Long
    Dim rowNum As Long
    
    'Define the path to the closed workbook
    Dim strFilePath As String
    strFilePath = "C:\VBA\GC_FISCALITÉ\DataConversion\CAR.xlsx"
    Dim strSheetName As String
    strSheetName = "CAR$"
    Dim strRange As String
    strRange = "A1:G120" 'Adjust the range as needed
    
    'Connection string for Excel
    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & strFilePath & ";" & _
                    "Extended Properties=""Excel 12.0 Xml;HDR=Yes"";"
    
    'Create a new ADO connection and recordset
    Dim cnn As Object: Set cnn = CreateObject("ADODB.Connection")
    Dim rst As Object: Set rst = CreateObject("ADODB.Recordset")
    
    'Open the connection
    cnn.Open strConnection
    
    'Open the recordset
    rst.Open "SELECT * FROM [" & strSheetName & strRange & "]", cnn, 3, 1, 1
    
    'Define the destination worksheet
    Set wsDest = wshFAC_Comptes_Clients
    
    'Get the last row in the destination sheet
    lastUsedRow = wsDest.Cells(wsDest.Rows.count, "A").End(xlUp).row
    rowNum = lastUsedRow
    
    'Loop through the recordset and write data to the destination sheet
    Dim client As String
    Dim dateFact As String
    Dim dateDue As String
    Dim factNo As String
    Dim clientCode As String
    Dim clientCodeFromDB As String
    Dim totalFact As Double
    Dim recu As Currency, regul As Currency
    Dim dateRecu As String
    Dim solde As Currency
    Dim joursDue As Long
    
    Dim errorMesg As String
    Dim totCAR As Double
    
    Do Until rst.EOF
        client = rst.Fields(0).value
        dateFact = rst.Fields(1).value
        factNo = rst.Fields(2).value
        totalFact = rst.Fields(3).value
        recu = rst.Fields(4).value
        regul = 0
        If IsNull(rst.Fields(5).value) Then
            dateRecu = ""
        Else
            dateRecu = rst.Fields(5).value
        End If
        solde = rst.Fields(6).value
        
        clientCode = Left(client, 10)
            clientCode = Left(clientCode, InStr(clientCode, " -") - 1)
        client = Mid(client, InStr(client, " - ") + 3, Len(client))
        totCAR = totCAR + solde
        
        'Is this a Valid Client ?
        Dim myInfo() As Variant
        Dim rng As Range: Set rng = wshBD_Clients.Range("dnrClients_Names_Only")
        myInfo = Fn_Find_Data_In_A_Range(rng, 1, client, 2)
        If myInfo(1) = "" Then
            If InStr(errorMesg, client) = 0 Then
                errorMesg = errorMesg & clientCode & " - " & client & vbNewLine
            End If
        End If
        clientCodeFromDB = myInfo(3)
        
        If clientCode <> clientCodeFromDB Then
            errorMesg = errorMesg & clientCode & " vs. " & clientCodeFromDB & vbNewLine
        End If
        
        wsDest.Cells(rowNum, fFacCCInvNo).value = factNo
        wsDest.Cells(rowNum, fFacCCInvoiceDate).value = dateFact
        wsDest.Cells(rowNum, fFacCCCustomer).value = client
        wsDest.Cells(rowNum, fFacCCCodeClient).value = clientCode
        wsDest.Cells(rowNum, fFacCCStatus).value = "Unpaid"
        wsDest.Cells(rowNum, fFacCCTerms).value = "Net"
        dateDue = DateAdd("d", 30, CDate(dateFact))
        wsDest.Cells(rowNum, fFacCCDueDate).value = Format$(CDate(dateDue), "mm/dd/yyyy")
        wsDest.Cells(rowNum, fFacCCTotal).value = totalFact
        wsDest.Cells(rowNum, fFacCCTotalPaid).value = recu
        wsDest.Cells(rowNum, fFacCCTotalRegul).value = 0
        wsDest.Cells(rowNum, fFacCCBalance).value = totalFact - recu + regul
        joursDue = DateDiff("d", dateDue, Date)
        wsDest.Cells(rowNum, fFacCCDaysOverdue).value = joursDue
        rowNum = rowNum + 1

        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les CAR ont été importés, pour un total de " & Format$(totCAR, "#,##0.00$")
    End If
    
    'Libérer la mémoire
    rst.Close
    cnn.Close
    Set rng = Nothing
    Set rst = Nothing
    Set cnn = Nothing
    Set wsDest = Nothing
    
End Sub

Sub Compare2ExcelFiles() '------------------------------------------ 2024-09-02 @ 06:24
    
    Application.ScreenUpdating = False
    
    'Declare and open the 2 workbooks
    Dim wbWas As Workbook
    Set wbWas = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx", ReadOnly:=True)
    Debug.Print "#066 - " & wbWas.Name
    Dim wbNow As Workbook
    Set wbNow = Workbooks.Open("C:\VBA\GC_FISCALITÉ\GCF_DataFiles\2024_09_01_1835\GCF_BD_Entrée_TBA.xlsx", ReadOnly:=True)
    Debug.Print "#067 - " & wbNow.Name

    'Declare the 2 worksheets
    Dim wsWas As Worksheet
    Set wsWas = wbWas.Worksheets("Clients")
    Dim wsNow As Worksheet
    Set wsNow = wbNow.Worksheets("Clients")
    
    'Détermine la dernière ligne utilisée dans chacune des 2 feuilles
    Dim lastUsedRowWas As Long
    lastUsedRowWas = wsWas.Cells(wsWas.Rows.count, 1).End(xlUp).row
    Dim lastUsedRowNOw As Long
    lastUsedRowNOw = wsNow.Cells(wsNow.Rows.count, 1).End(xlUp).row
    
    'Détermine le nombre de colonnes dans l'ancienne feuille
    Dim lastUsedColWas As Long
    lastUsedColWas = wsWas.Cells(wsWas.Columns.count).End(xlToLeft).Column
    
    'Erase and create a new worksheet for differences
    Dim wsNameStr As String
    wsNameStr = "X_Différences"
    Dim wsDiff As Worksheet
    Call CreateOrReplaceWorksheet(wsNameStr)
    Set wsDiff = ThisWorkbook.Worksheets(wsNameStr)
    wsDiff.Range("A1").value = "Ligne"
    wsDiff.Range("B1").value = "Colonne"
    wsDiff.Range("C1").value = "CodeClient"
    wsDiff.Range("D1").value = "Nom du Client"
    wsDiff.Range("E1").value = "Avant changement"
    wsDiff.Range("F1").value = "Type"
    wsDiff.Range("G1").value = "Après changement"
    Call Make_It_As_Header(wsDiff.Range("A1:G1"))

    Dim diffRow As Long
    diffRow = 2                                  'Take into consideration the Header
    Dim diffCol As Long
    diffCol = 1

    'Parcourir chaque ligne de l'ancienne version
    Dim cellWas As Range, cellNow As Range
    Dim foundRow As Range
    Dim clientCode As String
    Dim differences As String
    Dim readCells As Long
    Dim i As Long, j As Long
    For i = 1 To lastUsedRowWas
        clientCode = CStr(wsWas.Cells(i, 2).value)
        'Trouver la ligne correspondante dans la nouvelle version
        Set foundRow = wsNow.Columns(2).Find(What:=clientCode, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundRow Is Nothing Then
            Debug.Print "#068 - Ligne : " & i
            'Comparer les cellules des lignes correspondantes
            For j = 1 To lastUsedColWas
                readCells = readCells + 1
                Set cellWas = wsWas.Cells(i, j)
                Set cellNow = wsNow.Cells(foundRow.row, j)
                If CStr(cellWas.value) <> CStr(cellNow.value) Then
                    wsDiff.Cells(diffRow, 1).value = i
                    wsDiff.Cells(diffRow, 2).value = j
                    wsDiff.Cells(diffRow, 3).value = wsWas.Cells(i, 2).value
                    wsDiff.Cells(diffRow, 4).value = wsWas.Cells(cellWas.row, 1).value
                    wsDiff.Cells(diffRow, 5).value = cellWas.value
                    wsDiff.Cells(diffRow, 6).value = "'--->"
                    wsDiff.Cells(diffRow, 7).value = cellNow.value
                    diffRow = diffRow + 1
                End If
            Next j
        Else
            wsDiff.Cells(diffRow, 1).value = i
            wsDiff.Cells(diffRow, 3).value = wsWas.Cells(i, 2).value
            wsDiff.Cells(diffRow, 4).value = wsWas.Cells(cellWas.row, 1).value
            wsDiff.Cells(diffRow, 5).value = cellWas.value
            wsDiff.Cells(diffRow, 6).value = "XXXX"
            diffRow = diffRow + 1
        End If
    Next i
            
    wsDiff.Columns.AutoFit
    
    'Result print setup - 2024-08-05 @ 05:16
    diffRow = diffRow + 1
    wsDiff.Range("A" & diffRow).value = "**** " & Format$(readCells, "###,##0") & _
                                        " cellules analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:G" & diffRow)
    Call modAppli_Utils.ApplyConditionalFormatting(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:DC" & diffRow)
    Dim header1 As String: header1 = "Vérification des différences"
    Dim header2 As String: header2 = "Clients"
    Call Simple_Print_Setup(wsDiff, rngToPrint, header1, header2, "$1:$1", "P")
    
    Application.ScreenUpdating = True
    
    wsDiff.Activate

    'Close the workbooks without saving
    wbWas.Close SaveChanges:=False
    wbNow.Close SaveChanges:=False
    
    'Libérer la mémoire
    Set cellWas = Nothing
    Set cellNow = Nothing
    Set foundRow = Nothing
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set wbWas = Nothing
    Set wbNow = Nothing
    Set wsWas = Nothing
    Set wsNow = Nothing
    Set wsDiff = Nothing
    
    MsgBox "La comparaison est complétée." & vbNewLine & vbNewLine & _
           differences, vbInformation
           
End Sub

Sub AdjustClientNameInTEC()  '2024-08-03 @ 09:40

    'Définir les chemins d'accès des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Master.xlsx"
    Dim clientMF As String
    clientMF = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("TEC_Local")
    
    'Détermine la dernière rangée utilisée dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.Rows.count, 1).End(xlUp).row
    Dim lastUsedCol As Long
    lastUsedCol = sourceSheet.Cells(1, sourceSheet.Columns.count).End(xlToLeft).Column
    
    'Define the range to copy
    Dim sourceRange As Range
    Set sourceRange = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastUsedRow, lastUsedCol))
    
    ' Open the destination workbook
    Dim referenceWorkbook As Workbook: Set referenceWorkbook = Workbooks.Open(clientMF)
    Dim referenceSheet As Worksheet: Set referenceSheet = referenceWorkbook.Worksheets("Clients")
    Dim lastUsedRowClient As Long
    lastUsedRowClient = referenceSheet.Cells(referenceSheet.Rows.count, "A").End(xlUp).row
    
    Dim dictClients As Dictionary 'Code, Nom du Client
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowClient
        dictClients.Add CStr(referenceSheet.Cells(i, 2).value), referenceSheet.Cells(i, 1).value
    Next i
    
    Dim codeClient As String, nomClient As String, updatedNomClient As String
    For i = 2 To lastUsedRow
        codeClient = sourceSheet.Cells(i, 5).value
        nomClient = sourceSheet.Cells(i, 6).value
        updatedNomClient = dictClients(codeClient)
        Debug.Print "#069 - " & i & " : " & codeClient & " - " & nomClient & " ---> " & updatedNomClient
        sourceSheet.Cells(i, 6).value = updatedNomClient
    Next i
    
    'Save and close the destination workbook
    sourceWorkbook.Save
    sourceWorkbook.Close
    
    'Libérer la mémoire
    Set dictClients = Nothing
    Set sourceSheet = Nothing
    Set sourceRange = Nothing
    Set sourceWorkbook = Nothing
    Set referenceSheet = Nothing
    Set referenceWorkbook = Nothing
    
    MsgBox "Les données ont été copiées avec succès dans le fichier destination."
    
End Sub

Sub AdjustClientNameInCAR()  '2024-08-07 @ 17:11

    Dim sourceRange As Range
    
    'Définir les chemins d'accès des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx"
    Dim clientMF As String
    clientMF = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("CAR")
    
    'Détermine la dernière rangée utilisée dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.Rows.count, 1).End(xlUp).row
    Dim lastUsedCol As Long
    lastUsedCol = sourceSheet.Cells(1, sourceSheet.Columns.count).End(xlToLeft).Column
    
    'Define the range to copy
    Set sourceRange = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastUsedRow, lastUsedCol))
    
    ' Open the destination workbook
    Dim referenceWorkbook As Workbook: Set referenceWorkbook = Workbooks.Open(clientMF)
    Dim referenceSheet As Worksheet: Set referenceSheet = referenceWorkbook.Worksheets("Clients")
    Dim lastUsedRowClient As Long
    lastUsedRowClient = referenceSheet.Range("A9999").End(xlUp).row
    
    Dim dictClients As Dictionary 'Code, Nomdu Client
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowClient
        dictClients.Add CStr(referenceSheet.Cells(i, 2).value), referenceSheet.Cells(i, 1).value
'        Debug.Print "#070 - " & referenceSheet.Cells(i, 2).Value & " - " & referenceSheet.Cells(i, 1).Value
    Next i
    
    Dim codeClient As String, nomClient As String, updatedNomClient As String
    For i = 3 To lastUsedRow
        codeClient = sourceSheet.Cells(i, 4).value
        nomClient = sourceSheet.Cells(i, 3).value
        updatedNomClient = dictClients(codeClient)
        Debug.Print "#071 - " & i & " : " & codeClient & " - " & nomClient & " ---> " & updatedNomClient
        sourceSheet.Cells(i, 3).value = updatedNomClient
    Next i
    
    'Save and close the destination workbook
    sourceWorkbook.Save
    sourceWorkbook.Close
    
    'Libérer la mémoire
    Set dictClients = Nothing
    Set sourceSheet = Nothing
    Set sourceRange = Nothing
    Set sourceWorkbook = Nothing
    Set referenceSheet = Nothing
    Set referenceWorkbook = Nothing
    
    MsgBox "Les données ont été copiées avec succès dans le fichier destination."
    
End Sub

Sub CheckClientName() '2024-08-10 @ 10:13

    'Définir les chemins d'accès des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("Clients")
    
    'Détermine la dernière rangée utilisée dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.Rows.count, 1).End(xlUp).row
    
    Dim codeClient As String, nomClient As String, contactFact As String
    Dim i As Long
    For i = 2 To lastUsedRow
        codeClient = sourceSheet.Cells(i, fClntFMClientID).value
        nomClient = Trim(sourceSheet.Cells(i, fClntFMClientNom).value)
        contactFact = Trim(sourceSheet.Cells(i, fClntFMContactFacturation).value)
        If InStr(nomClient, contactFact) = 0 Then
            Debug.Print "#072 - " & i & " : " & codeClient & " - " & nomClient & " on ajoute '" & contactFact & "'"
        End If
    Next i
    
    'Save and close the destination workbook
    sourceWorkbook.Save
    sourceWorkbook.Close
    
    'Libérer la mémoire
    Set sourceSheet = Nothing
    Set sourceWorkbook = Nothing
    
    MsgBox "Les données ont été vérifiées avec succès dans le fichier Clients."
    
End Sub

Sub ConstruireSommaireHeures() '2024-08-12 @ 21:09

    'Définir les chemins d'accès des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALITÉ\GCF_DataFiles\GCF_BD_MASTER.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("TEC_Local")
    
    'Détermine la dernière rangée utilisée dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.Rows.count, 1).End(xlUp).row
    
    Dim profID As Long
    Dim prof As String, codeClient As String, nomClient As String
    Dim estFacturable As String, estFacturee As String, estDetruit As String
    Dim dateServ As Date
    Dim hresSaisies As Double, hresDetruites As Double, hresFacturees As Double
    Dim hresNonDetruites As Double, hresFacturables As Double, hresNonFacturables As Double
    Dim i As Long
    For i = 2 To lastUsedRow
        profID = sourceSheet.Cells(i, 2).value
        prof = sourceSheet.Cells(i, 3).value
        dateServ = sourceSheet.Cells(i, 4).value
        codeClient = sourceSheet.Cells(i, 5).value
        nomClient = Trim(sourceSheet.Cells(i, 6).value)
        hresSaisies = Trim(sourceSheet.Cells(i, 8).value)
        estFacturable = sourceSheet.Cells(i, 10).value
        estFacturee = sourceSheet.Cells(i, 12).value
        estDetruit = sourceSheet.Cells(i, 14).value
        
        hresDetruites = 0
        If estDetruit = "VRAI" Then
            hresDetruites = hresSaisies
        End If
        hresNonDetruites = hresSaisies - hresDetruites
        
        hresFacturables = 0
        hresNonFacturables = 0
        If estFacturable = "VRAI" Then
            hresFacturables = hresNonDetruites
        Else
            hresNonFacturables = hresNonDetruites
        End If
        
        hresFacturees = 0
        If estFacturee = "VRAI" Then
            hresFacturees = hresNonDetruites
        End If
        
    Next i
    
    'Close the source workbook
    sourceWorkbook.Close
    
    'Libérer la mémoire
    Set sourceSheet = Nothing
    Set sourceWorkbook = Nothing
    
    MsgBox "Sommaire des heures est complété."
    
End Sub

Sub CorrigeNomClientInTEC()  '2024-08-23 @ 06:32

    'Source - Définir les chemins d'accès des fichiers, le Workbook, le Worksheet et le Range
    Dim sourceFilePath As String
    sourceFilePath = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Master.xlsx"
    Dim wbSource As Workbook: Set wbSource = Workbooks.Open(sourceFilePath)
    Dim wsSource As Worksheet: Set wsSource = wbSource.Worksheets("TEC_Local")
    
    'Détermine la dernière rangée et dernière colonne utilisées dans wshTEC_Local
    Dim lastUsedRowTEC As Long
    lastUsedRowTEC = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
    
    'Open the Master File Workbook
    Dim clientMFPath As String
    clientMFPath = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx"
    Dim wbMF As Workbook: Set wbMF = Workbooks.Open(clientMFPath)
    Dim wsMF As Worksheet: Set wsMF = wbMF.Worksheets("Clients")
    Dim lastUsedRowTECClient As Long
    lastUsedRowTECClient = wsMF.Cells(wsMF.Rows.count, 1).End(xlUp).row
    
    'Setup output file
    Dim strOutput As String
    strOutput = "X_TEC_Correction_Nom"
    Call CreateOrReplaceWorksheet(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").value = "TEC_Nom_Client"
    wsOutput.Range("B1").value = "Code_de_Client"
    wsOutput.Range("C1").value = "Nom_Client_Master"
    wsOutput.Range("D1").value = "TECID"
    wsOutput.Range("E1").value = "TEC_Prof"
    wsOutput.Range("F1").value = "TEC_Date"
    Call Make_It_As_Header(wsOutput.Range("A1:F1"))
    
    'Build the dictionnary (Code, Nom du client) from Client's Master File
    Dim dictClients As Dictionary
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowTECClient
        dictClients.Add CStr(wsMF.Cells(i, 2).value), wsMF.Cells(i, 1).value
    Next i
    
    'Parse TEC_Local to verify TEC's clientName vs. MasterFile's clientName
    Dim codeClientTEC As String, nomClientTEC As String, nomClientFromMF As String
    Dim casDelta As Long, rowOutput As Long
    rowOutput = 2
    For i = 2 To lastUsedRowTEC
        codeClientTEC = wsSource.Cells(i, 5).value
        nomClientTEC = wsSource.Cells(i, 6).value
        nomClientFromMF = dictClients(codeClientTEC)
        If nomClientTEC <> nomClientFromMF Then
            Debug.Print "#073 - " & i & " : " & codeClientTEC & " - " & nomClientTEC & " <---> " & nomClientFromMF
            wsSource.Cells(i, 6).value = nomClientFromMF
            wsOutput.Cells(rowOutput, 1).value = nomClientTEC
            wsOutput.Cells(rowOutput, 2).value = codeClientTEC
            wsOutput.Cells(rowOutput, 3).value = nomClientFromMF
            wsOutput.Cells(rowOutput, 4).value = wsSource.Cells(i, 1).value
            wsOutput.Cells(rowOutput, 5).value = wsSource.Cells(i, 3).value
            wsOutput.Cells(rowOutput, 6).value = wsSource.Cells(i, 4).value
            rowOutput = rowOutput + 1
            casDelta = casDelta + 1
        End If
    Next i
    
    wsOutput.Columns.AutoFit

    'Result print setup - 2024-08-05 @ 05:16
    rowOutput = rowOutput + 1
    wsOutput.Range("A" & rowOutput).value = "**** " & Format$(lastUsedRowTEC - 1, "###,##0") & _
                                        " lignes analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("A2:F" & rowOutput)
    Call modAppli_Utils.ApplyConditionalFormatting(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:E" & rowOutput)
    Dim header1 As String: header1 = "Correction des noms de clients dans les TEC"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wbSource.Close SaveChanges:=True
    wbMF.Close SaveChanges:=False

    'Libérer la mémoire
    Set dictClients = Nothing
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set wsSource = Nothing
    Set wsOutput = Nothing
    Set wbSource = Nothing
    Set wsMF = Nothing
    Set wbMF = Nothing
    
    MsgBox "Il y a " & casDelta & " cas où le nom du client (TEC) diffère" & _
            vbNewLine & vbNewLine & "du nom de client du Fichier MAÎTRE", vbInformation
    
End Sub

Public Sub CorrigeNomClientInCAR()  '2024-08-31 @ 06:52

    'Worksheets to be corrected - Open the workbook (worksheet will be determined later)
    Dim sourceFilePath As String
    sourceFilePath = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Master.xlsx"
    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(sourceFilePath)
    
    'Client's Master File - Workbook & Worksheet
    Dim clientMFPath As String
    clientMFPath = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx"
    Dim wbMF As Workbook
    Set wbMF = Workbooks.Open(clientMFPath)
    Dim wsMF As Worksheet
    Set wsMF = wbMF.Worksheets("Clients")
    Dim lastUsedRowClientMF As Long
    lastUsedRowClientMF = wsMF.Cells(wsMF.Rows.count, 1).End(xlUp).row
    
    'Build the dictionnary (Code, Nom du client) from Client's Master File
    Dim clientName As String
    Dim dictClients As Dictionary
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowClientMF
        'Enlève les informations de contact
        clientName = wsMF.Cells(i, fClntFMClientNom).value
        Do While InStr(clientName, "[") > 0 And InStr(clientName, "]") > 0
            clientName = Fn_Strip_Contact_From_Client_Name(clientName)
        Loop
        dictClients.Add CStr(wsMF.Cells(i, fClntFMClientID).value), clientName
    Next i
    
    'Setup output file
    Dim strOutput As String
    strOutput = "X_CAR_Correction_Nom"
    Call CreateOrReplaceWorksheet(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").value = "Feuille"
    wsOutput.Range("B1").value = "No Facture"
    wsOutput.Range("C1").value = "Nom de client (Facture)"
    wsOutput.Range("D1").value = "Code_de_Client"
    wsOutput.Range("E1").value = "Changé pour"
    Call Make_It_As_Header(wsOutput.Range("A1:E1"))
    Dim rowOutput As Long
    rowOutput = 2 'Skip the header
    
    'There is 2 worksheets to adjust
    Dim ws As Variant
    Dim wsSource As Worksheet
    Dim casDelta As Long
    For Each ws In Array("FAC_Entête|4|6", "FAC_Comptes_Clients|4|3")
        Dim param As String
        param = Mid(ws, InStr(ws, "|") + 1)
        ws = Left(ws, InStr(ws, "|") - 1)
        Dim colClientID As Long, colClientName As Long
        colClientID = Left(param, InStr(param, "|") - 1)
        colClientName = Right(param, Len(param) - InStr(param, "|"))
        'Set the worksheet Object
        Set wsSource = wbSource.Sheets(ws)
        'Détermine la dernière rangée utilisée dans la feuille
        Dim lastUsedRow As Long
        lastUsedRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
        Dim codeClientCAR As String, nomClientCAR As String, nomClientFromMF As String
        For i = 3 To lastUsedRow
            codeClientCAR = wsSource.Cells(i, colClientID).value
            nomClientCAR = wsSource.Cells(i, colClientName).value
            nomClientFromMF = dictClients(codeClientCAR)
            If nomClientCAR <> nomClientFromMF Then
                Debug.Print "#074 - " & i & " : " & codeClientCAR & " - " & nomClientCAR & " <---> " & nomClientFromMF
                wsSource.Cells(i, colClientName).value = nomClientFromMF
                wsOutput.Cells(rowOutput, 1).value = wsSource.Name
                wsOutput.Cells(rowOutput, 2).value = wsSource.Cells(i, 1).value
                wsOutput.Cells(rowOutput, 3).value = nomClientCAR
                wsOutput.Cells(rowOutput, 4).value = codeClientCAR
                wsOutput.Cells(rowOutput, 5).value = nomClientFromMF
                rowOutput = rowOutput + 1
                casDelta = casDelta + 1
            End If
        Next i
    Next ws
    
    wsOutput.Columns.AutoFit

    'Result print setup - 2024-08-05 @ 05:16
    rowOutput = rowOutput + 1
    wsOutput.Range("A" & rowOutput).value = "**** " & Format$(lastUsedRow - 1, "###,##0") & _
                                        " lignes analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("A2:F" & rowOutput)
    Call modAppli_Utils.ApplyConditionalFormatting(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:E" & rowOutput)
    Dim header1 As String: header1 = "Correction des noms de clients dans les CAR"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wbSource.Close SaveChanges:=True
    wbMF.Close SaveChanges:=False

    'Libérer la mémoire
    Set dictClients = Nothing
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set wsOutput = Nothing
    Set wsSource = Nothing
    Set wbSource = Nothing
    Set ws = Nothing
    Set wsMF = Nothing
    Set wbMF = Nothing
    
''    MsgBox "Il y a " & casDelta & " cas où le nom du client (TEC) diffère" & _
''            vbNewLine & vbNewLine & "du nom de client du Fichier MAÎTRE", vbInformation
'
End Sub

Sub ImporterDonnéesManquantes_CAR() '2024-08-24 @ 15:58

    Application.ScreenUpdating = False
    
    'Declare and open the 2 workbooks
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER - Copie.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx")

    'Declare the 2 worksheets
    Dim ws1 As Worksheet: Set ws1 = wb1.Worksheets("FAC_Comptes_Clients")
    Dim ws2 As Worksheet: Set ws2 = wb2.Worksheets("FAC_Entête")
    
    Dim lastUsedRow As Long
    lastUsedRow = ws1.Cells(ws1.Rows.count, 1).End(xlUp).row
    Dim row As Long, rowFAC_Entete As Long
    row = 2
    rowFAC_Entete = ws2.Cells(ws2.Rows.count, 1).End(xlUp).row + 1
    
    Dim i As Integer
    For i = 2 To lastUsedRow
        If InStr(ws1.Range("A" & i).value, "24-") <> 1 Then
            ws2.Range("A" & rowFAC_Entete).value = ws1.Range("A" & i)
            ws2.Range("B" & rowFAC_Entete).value = ws1.Range("B" & i)
            ws2.Range("C" & rowFAC_Entete).value = "C"
            ws2.Range("D" & rowFAC_Entete).value = ws1.Range("D" & i)
            ws2.Range("F" & rowFAC_Entete).value = ws1.Range("C" & i)
            
            ws2.Range("J" & rowFAC_Entete).value = ws1.Range("H" & i)
            ws2.Range("K" & rowFAC_Entete).value = "Frais de poste"
            ws2.Range("L" & rowFAC_Entete).value = 0
            ws2.Range("M" & rowFAC_Entete).value = "Frais d'expert en taxes"
            ws2.Range("N" & rowFAC_Entete).value = 0
            ws2.Range("O" & rowFAC_Entete).value = "Autres frais"
            ws2.Range("P" & rowFAC_Entete).value = 0
            
            ws2.Range("Q" & rowFAC_Entete).value = Format$(CDbl(5), "#0.000")
            ws2.Range("R" & rowFAC_Entete).value = 0
            ws2.Range("S" & rowFAC_Entete).value = Format$(CDbl(9.975), "#0.000")
            ws2.Range("T" & rowFAC_Entete).value = 0
            
            ws2.Range("U" & rowFAC_Entete).value = ws1.Range("H" & i)
            ws2.Range("V" & rowFAC_Entete).value = 0
            rowFAC_Entete = rowFAC_Entete + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    'Close the workbooks without saving
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=True
    
    'Libérer la mémoire
    Set wb1 = Nothing
    Set wb2 = Nothing
    Set ws1 = Nothing
    Set ws2 = Nothing
    
    MsgBox "Le traitement est complété", vbInformation
           
End Sub

Sub FusionnerDonnéesManquantes_CAR() '2024-08-29 @ 07:29

    Application.ScreenUpdating = False
    
    'Declare and open the 2 workbooks
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\GCF_DataFiles\CAR_A_COMPLÉTER.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx")

    'Declare the 2 worksheets
    Dim ws1 As Worksheet: Set ws1 = wb1.Worksheets("Feuil1")
    Dim ws2 As Worksheet: Set ws2 = wb2.Worksheets("FAC_Entête")
    
    Dim lastUsedRow As Long
    lastUsedRow = ws1.Cells(ws1.Rows.count, 1).End(xlUp).row
    Dim lastUsedRowTarget As Long
    lastUsedRowTarget = ws2.Cells(ws2.Rows.count, 1).End(xlUp).row
    
    'Define the target Range
    Dim rngTarget As Range: Set rngTarget = ws2.Range("A2:A" & lastUsedRowTarget)
    
    Dim invNo As String
    Dim hono As Currency, af1 As Currency, af2 As Currency, af3 As Currency
    Dim tps As Currency, tvq As Currency, arTotal As Currency, depot As Currency
    Dim af1Str As String, af2Str As String, af3Str As String
    Dim foundCells As Range
    Dim t() As Currency
    ReDim t(1 To 8)

    Dim i As Integer, ii As Integer
    
    For i = 2 To lastUsedRow
        invNo = ws1.Cells(i, 1).value
        
        hono = ws1.Cells(i, 10).value
        af1Str = ws1.Cells(i, 11).value
        af1 = ws1.Cells(i, 12).value
        af2Str = ws1.Cells(i, 13).value
        af2 = ws1.Cells(i, 14).value
        af3Str = ws1.Cells(i, 15).value
        af3 = ws1.Cells(i, 16).value
        tps = ws1.Cells(i, 18).value
        tvq = ws1.Cells(i, 20).value
        arTotal = ws1.Cells(i, 21).value
        depot = ws1.Cells(i, 22).value

        If hono + af1 + af2 + af3 + tps + tvq <> arTotal Then
            MsgBox "Une ligne (" & i & ") ne balance pas !!!", vbCritical
        End If
        
        'Find the InvNo in wshFAC_Entête
        Set foundCells = rngTarget.Columns(1).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
        If foundCells Is Nothing Then
            MsgBox "**** Je n'ai pas trouvé la facture '" & invNo & "' dans wshFAC_Entête", vbCritical
        Else
            ii = foundCells.row
        End If
        
        If ws2.Cells(ii, 21).value <> arTotal Then
            MsgBox "Problème d'intégrité pour la facture '" & invNo & "' au niveau de arTotal", vbCritical
        End If
        
        'Replace values in Target, with the Source info
        ws2.Cells(ii, 10).value = hono
        If af1 <> 0 And af1Str <> ws2.Cells(ii, 11) Then
            ws2.Cells(ii, 11).value = af1Str
        End If
        ws2.Cells(ii, 12).value = af1
        
        If af2 <> 0 And af2Str <> ws2.Cells(ii, 13) Then
            ws2.Cells(ii, 13).value = af2Str
        End If
        ws2.Cells(ii, 14).value = af2
        
        If af3 <> 0 And af3Str <> ws2.Cells(ii, 15) Then
            ws2.Cells(ii, 15).value = af3Str
        End If
        ws2.Cells(ii, 16).value = af3
        
        ws2.Cells(ii, 18).value = tps
        ws2.Cells(ii, 20).value = tvq
        
        ws2.Cells(ii, 22).value = depot
        
        If ws2.Cells(ii, 10) + ws2.Cells(ii, 12) + ws2.Cells(ii, 14) + ws2.Cells(ii, 16) + _
            ws2.Cells(ii, 18) + ws2.Cells(ii, 20) <> ws2.Cells(ii, 21).value Then
            MsgBox "Problème avec les assignations...", vbCritical
        End If
        
        t(1) = t(1) + ws2.Cells(ii, 10)
        t(2) = t(2) + ws2.Cells(ii, 12)
        t(3) = t(3) + ws2.Cells(ii, 14)
        t(4) = t(4) + ws2.Cells(ii, 16)
        t(5) = t(5) + ws2.Cells(ii, 18)
        t(6) = t(6) + ws2.Cells(ii, 20)
        t(7) = t(7) + ws2.Cells(ii, 21)
        t(8) = t(8) + ws2.Cells(ii, 22)
        
        Debug.Print "#075 - " & "x8", invNo, ii, Format(i / lastUsedRow, "##0.00 %")
    
    Next i
    
    Debug.Print "#076 - " & t(1), t(2), t(3), t(4), t(5), t(6), t(7), t(8)
    
    Application.ScreenUpdating = True
    
    'Close the workbooks without saving
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=True
    
    'Libérer la mémoire
    Set foundCells = Nothing
    Set rngTarget = Nothing
    Set wb1 = Nothing
    Set wb2 = Nothing
    Set ws1 = Nothing
    Set ws2 = Nothing
    
    MsgBox "Le traitement est complété", vbInformation
           
End Sub


