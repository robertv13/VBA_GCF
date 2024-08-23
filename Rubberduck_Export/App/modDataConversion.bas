Attribute VB_Name = "modDataConversion"
Option Explicit

'Importation des clients � partir de ... \DataConversion\Clients.xlsx
Sub Copy_Data_Between_Closed_Workbooks_Clients() '2024-08-03 @ 09:40

    Dim sourceRange As Range
    
    'D�finir les chemins d'acc�s des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\DataConversion\Clients.xlsx"
    Dim destinationFilePath As String
    destinationFilePath = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & "GCF_BD_Entr�e.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("Feuil1")
    
    'D�termine la derni�re rang�e utilis�e dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.rows.count, 1).End(xlUp).row
    Dim lastUsedCol As Long
    lastUsedCol = sourceSheet.Cells(1, sourceSheet.columns.count).End(xlToLeft).Column
    
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
    sourceWorkbook.Close saveChanges:=False
    
    'Clean up
    Set sourceSheet = Nothing
    Set sourceRange = Nothing
    Set sourceWorkbook = Nothing
    Set destinationSheet = Nothing
    Set destinationWorkbook = Nothing
    
    MsgBox "Les donn�es ont �t� copi�es avec succ�s dans le fichier destination."
    
End Sub

'Ajustements � la feuille DB_Clients (*) ---> [*]
Sub Clients_Ajuste_Nom()

    'Declare and open the closed workbook
    Dim wb As Workbook
    Set wb = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx")

    'Define the worksheet you want to work with
    Dim ws As Worksheet
    Set ws = wb.Worksheets("Clients")
    
    'Find the last used row with data in column A
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    
    'Loop through each row starting from row 2 (headers are 1 row)
    Dim client As String, client_ID As String, contactFacturation As String
    Dim posOpenParenthesis As Integer, posCloseParenthesis As Integer
    Dim numberOpenParenthesis As Integer, numberCloseParenthesis As Integer
    Dim i As Long
    For i = 2 To lastUsedRow
        'Load data into variables
        client = ws.Cells(i, 1).value
        client_ID = ws.Cells(i, 2).value
        contactFacturation = ws.Cells(i, 3).value
        
        'Process the data and make adjustments if necessary
        posOpenParenthesis = InStr(client, "(")
        posCloseParenthesis = InStr(client, ")")
        numberOpenParenthesis = CountCharOccurrences(client, "(")
        numberCloseParenthesis = CountCharOccurrences(client, ")")
        
        If numberOpenParenthesis = 1 And numberCloseParenthesis = 1 Then
            If posCloseParenthesis > posOpenParenthesis + 5 Then
                client = Replace(client, "(", "[")
                client = Replace(client, ")", "]")
                ws.Cells(i, 1).value = client
                Debug.Print i & " - " & client
            End If
        End If
        
    Next i
    
    wb.Save
    
    MsgBox "Le traitement est compl�t� sur " & i - 1 & " lignes"
    
End Sub

'Ajustements � la feuille DB_Clients (Ajout du contactdans le nom du client)
Sub Clients_Ajout_Contact_Dans_Nom()

    'Declare and open the closed workbook
    Dim wb As Workbook
    Set wb = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx")

    'Define the worksheet you want to work with
    Dim ws As Worksheet
    Set ws = wb.Worksheets("Clients")
    
    'Find the last used row with data in column A
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    
    'Loop through each row starting from row 2 (headers are 1 row)
    Dim client As String, client_ID As String, contactFacturation As String
    Dim posOpenSquareBracket As Integer, posCloseSquareBracket As Integer
    Dim numberOpenSquareBracket As Integer, numberCloseSquareBracket As Integer
    Dim i As Long
    For i = 2 To lastUsedRow
        'Load data into variables
        client = ws.Cells(i, 1).value
        client_ID = ws.Cells(i, 2).value
        contactFacturation = Trim(ws.Cells(i, 3).value)
        
        'Process the data and make adjustments if necessary
        posOpenSquareBracket = InStr(client, "[")
        posCloseSquareBracket = InStr(client, "]")
        
        If numberOpenSquareBracket = 0 And numberCloseSquareBracket = 0 Then
            If contactFacturation <> "" And InStr(client, contactFacturation) = 0 Then
                client = Trim(client) & " [" & contactFacturation & "]"
                ws.Cells(i, 1).value = client
                Debug.Print i & " - " & client
            End If
        End If
        
    Next i
    
    wb.Save
    
    MsgBox "Le traitement est compl�t� sur " & i - 1 & " lignes"
    
End Sub

Sub Import_Data_From_Closed_Workbooks_TEC() '2024-08-14 @ 06:43 & 2024-08-03 @ 16:15

    Call Client_List_Import_All
    
    'Define the path to the closed workbook
    Dim strFilePath As String
    strFilePath = "C:\VBA\GC_FISCALIT�\DataConversion\TEC_20240814.xlsx"
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
    Set wsDest = ThisWorkbook.Sheets("TEC_Local")
    
    'Get the last row in the destination sheet
    Dim lastUsedRow As Long, rowNum As Long
    lastUsedRow = wsDest.Range("A999").End(xlUp).row
    rowNum = lastUsedRow
    
    'Loop through the recordset and write data to the destination sheet
    Dim prof As String
    Dim client As String
    Dim clientCode As String
    Dim clientCodeFromDB As String
    Dim errorMesg As String
    Dim TEC_ID As Long: TEC_ID = 342
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
        
        TEC_ID = TEC_ID + 1
        wsDest.Range("A" & rowNum).value = TEC_ID
        wsDest.Range("B" & rowNum).value = Get_ID_From_Initials(prof)
        wsDest.Range("C" & rowNum).value = prof
        wsDest.Range("D" & rowNum).value = rst.Fields(1).value
        wsDest.Range("E" & rowNum).value = clientCode
        wsDest.Range("F" & rowNum).value = client
        wsDest.Range("G" & rowNum).value = rst.Fields(4).value
        wsDest.Range("H" & rowNum).value = rst.Fields(5).value
        wsDest.Range("I" & rowNum).value = ""
        wsDest.Range("J" & rowNum).value = "VRAI"
        wsDest.Range("K" & rowNum).value = Format$(Now(), "dd/mm/yyyy hh:nn:ss")
        wsDest.Range("L" & rowNum).value = "FAUX"
        wsDest.Range("M" & rowNum).value = ""
        wsDest.Range("N" & rowNum).value = "FAUX"
        wsDest.Range("O" & rowNum).value = ThisWorkbook.name
        wsDest.Range("P" & rowNum).value = ""
        
        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les TEC ont �t� import�s, pour un total de " & totHres & " heures"
    End If
    
    'Clean up
    rst.Close
    cnn.Close
    
    Set rst = Nothing
    Set cnn = Nothing
    
End Sub

'Only valid for this conversion process
Function Get_ID_From_Initials(p As String) As Long

    Select Case p
        Case "GC"
            Get_ID_From_Initials = 1
        Case "VG"
            Get_ID_From_Initials = 2
        Case "AR"
            Get_ID_From_Initials = 3
        Case "ML"
            Get_ID_From_Initials = 4
        Case Else
            Get_ID_From_Initials = 0
    End Select

End Function

'Importation des fournisseurs � partir de ... \DataConversion\Fournisseurs.xlsx
Sub Import_Data_From_Closed_Workbooks_Fournisseurs() '2024-08-03 @ 18:10

    'D�finir les chemins d'acc�s des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\DataConversion\Fournisseurs.xlsx"
    Dim destinationFilePath As String
    destinationFilePath = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & "GCF_BD_Entr�e.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("Fournisseurs")
    
    'D�termine la derni�re rang�e utilis�e dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.rows.count, 1).End(xlUp).row
    Dim lastUsedCol As Long
    lastUsedCol = sourceSheet.Cells(1, sourceSheet.columns.count).End(xlToLeft).Column
    
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
    sourceWorkbook.Close saveChanges:=False
    
    'Clean up
    Set sourceSheet = Nothing
    Set sourceRange = Nothing
    Set sourceWorkbook = Nothing
    Set destinationSheet = Nothing
    Set destinationWorkbook = Nothing
    
    MsgBox "Les donn�es (fournisseurs) ont �t� copi�es avec succ�s dans" & vbNewLine & _
            vbNewLine & "le fichier destination."
    
End Sub

Sub Import_Data_From_Closed_Workbooks_GL_BV() '2024-08-03 @ 18:20

    'Define the path to the closed workbook
    Dim strFilePath As String
    strFilePath = "C:\VBA\GC_FISCALIT�\DataConversion\GL_BV.xlsx"
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
    Dim wsDest As Worksheet: Set wsDest = ThisWorkbook.Sheets("GL_Trans")
    
    'Get the last row in the destination sheet
    Dim lastUsedRow As Long
    lastUsedRow = wsDest.Range("A999").End(xlUp).row
    Dim rowNum As Long
    rowNum = lastUsedRow
    
    'Loop through the recordset and write data to the destination sheet
    Dim descriptionGL As String
    Dim codeGL As String
    Dim errorMesg As String
    Dim No_Entr�e As Long
    No_Entr�e = 1
    Dim amount As Double
    Dim totalDT As Double, totalCT As Double
    
    Do Until rst.EOF
        rowNum = rowNum + 1
        descriptionGL = rst.Fields(0).value
        codeGL = Fn_Get_GL_Code_From_GL_Description(descriptionGL)
        amount = rst.Fields(1).value
        If amount > 0 Then
            totalDT = totalDT + amount
        Else
            totalCT = totalCT - amount
        End If
        
        wsDest.Range("A" & rowNum).value = No_Entr�e
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
        wsDest.Range("J" & rowNum).value = Format$(Now(), "mm/dd/yyyy hh:mm:ss")
        
        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les soldes ont �t� import�s, pour un total d�bit de " & totalDT & _
                vbNewLine & vbNewLine & "un total cr�dit de " & totalCT
    End If
    
    'Clean up
    rst.Close
    cnn.Close
    
    Set rst = Nothing
    Set cnn = Nothing
    
End Sub

Sub Import_Data_From_Closed_Workbooks_CC() '2024-08-04 @ 07:31

    Call Client_List_Import_All
    
    Dim strConnection As String
    Dim wsDest As Worksheet
    Dim i As Long, j As Long
    Dim lastUsedRow As Long
    Dim rowNum As Long
    
    'Define the path to the closed workbook
    Dim strFilePath As String
    strFilePath = "C:\VBA\GC_FISCALIT�\DataConversion\CAR.xlsx"
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
    Set wsDest = ThisWorkbook.Sheets("CAR")
    
    'Get the last row in the destination sheet
    lastUsedRow = wsDest.Range("A999").End(xlUp).row
    rowNum = lastUsedRow
    
    'Loop through the recordset and write data to the destination sheet
    Dim client As String
    Dim dateFact As String
    Dim dateDue As String
    Dim factNo As String
    Dim clientCode As String
    Dim clientCodeFromDB As String
    Dim totalFact As Double
    Dim recu As Double
    Dim dateRecu As String
    Dim solde As Double
    Dim joursDue As Long
    
    Dim errorMesg As String
    Dim totCAR As Double
    
    Do Until rst.EOF
        client = rst.Fields(0).value
        dateFact = rst.Fields(1).value
        factNo = rst.Fields(2).value
        totalFact = rst.Fields(3).value
        recu = rst.Fields(4).value
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
        
        wsDest.Range("A" & rowNum).value = factNo
        wsDest.Range("B" & rowNum).value = dateFact
        wsDest.Range("C" & rowNum).value = client
        wsDest.Range("D" & rowNum).value = clientCode
        wsDest.Range("E" & rowNum).value = "Unpaid"
        wsDest.Range("F" & rowNum).value = "Net 30"
        dateDue = DateAdd("d", 30, CDate(dateFact))
        wsDest.Range("G" & rowNum).value = Format$(CDate(dateDue), "mm/dd/yyyy")
        wsDest.Range("H" & rowNum).value = totalFact
        wsDest.Range("I" & rowNum).value = recu
        wsDest.Range("J" & rowNum).value = totalFact - recu
        joursDue = DateDiff("d", dateDue, Date)
        wsDest.Range("K" & rowNum).value = joursDue
        rowNum = rowNum + 1

        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les CAR ont �t� import�s, pour un total de " & Format$(totCAR, "#,##0.00$")
    End If
    
    'Clean up
    rst.Close
    cnn.Close
    
    Set rst = Nothing
    Set cnn = Nothing
    
End Sub

Sub Compare_2_Excel_Files()                      '2024-08-05 @ 05:32

    Application.ScreenUpdating = False
    
    'Declare and open the 2 workbooks
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER_ACO.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER.xlsx")

    'Declare the 2 worksheets
    Dim ws1 As Worksheet
    Set ws1 = wb1.Worksheets("CAR")
    Dim ws2 As Worksheet
    Set ws2 = wb2.Worksheets("CAR")
    
    'Erase and create a new worksheet for differences
    Dim wsDiff As Worksheet
    Call CreateOrReplaceWorksheet("Diff�rences")
    Set wsDiff = ThisWorkbook.Worksheets("Diff�rences")
    wsDiff.Range("A1").value = "Position"
    wsDiff.Range("B1").value = "CodeClient"
    wsDiff.Range("C1").value = "Valeur originale"
    wsDiff.Range("D1").value = "Valeur corrig�e"
    Call Make_It_As_Header(wsDiff.Range("A1:D1"))

    Dim diffRow As Long
    diffRow = 2                                  'Take into consideration the Header
    Dim diffCol As Long
    diffCol = 1

    'Loop through each cell and compare
    Dim cell1 As Range
    Dim cell2 As Range
    Dim readCells As Long
    For Each cell1 In ws1.usedRange
        Set cell2 = ws2.Cells(cell1.row, cell1.Column)
        readCells = readCells + 1
        If cell1.value <> cell2.value Then
            wsDiff.Cells(diffRow, 1).value = "Ligne " & cell1.row & ", Colonne " & cell1.Column
            wsDiff.Cells(diffRow, 2).value = ws1.Cells(cell1.row, 4).value
            wsDiff.Cells(diffRow, 3).value = cell1.value
            wsDiff.Cells(diffRow, 4).value = cell2.value
            diffRow = diffRow + 1
        End If
    Next cell1

    wsDiff.columns.AutoFit
    
    'Result print setup - 2024-08-05 @ 05:16
    diffRow = diffRow + 1
    wsDiff.Range("A" & diffRow).value = "**** " & Format$(readCells, "###,##0") & _
                                        " cellules analys�es dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:D" & diffRow)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:DC" & diffRow)
    Dim header1 As String: header1 = "V�rification des diff�rences"
    Dim header2 As String: header2 = "Clients"
    Call Simple_Print_Setup(wsDiff, rngToPrint, header1, header2, "P")
    
    Application.ScreenUpdating = True
    
    wsDiff.Activate

    'Close the workbooks without saving
    wb1.Close False
    wb2.Close False
    
    'Cleanup
    Set cell1 = Nothing
    Set cell2 = Nothing
    Set rngToPrint = Nothing
    Set wb1 = Nothing
    Set wb2 = Nothing
    Set ws1 = Nothing
    Set ws2 = Nothing
    Set wsDiff = Nothing
    
    MsgBox "Comparison complete. " & vbNewLine & vbNewLine & _
           "Differences have been recorded in the 'Differences' sheet.", vbInformation
           
End Sub

Sub Adjust_Client_Name_In_TEC()  '2024-08-03 @ 09:40

    'D�finir les chemins d'acc�s des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Master.xlsx"
    Dim clientMF As String
    clientMF = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("TEC_Local")
    
    'D�termine la derni�re rang�e utilis�e dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.rows.count, 1).End(xlUp).row
    Dim lastUsedCol As Long
    lastUsedCol = sourceSheet.Cells(1, sourceSheet.columns.count).End(xlToLeft).Column
    
    'Define the range to copy
    Dim sourceRange As Range
    Set sourceRange = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastUsedRow, lastUsedCol))
    
    ' Open the destination workbook
    Dim referenceWorkbook As Workbook: Set referenceWorkbook = Workbooks.Open(clientMF)
    Dim referenceSheet As Worksheet: Set referenceSheet = referenceWorkbook.Worksheets("Clients")
    Dim lastUsedRowClient As Long
    lastUsedRowClient = referenceSheet.Range("A9999").End(xlUp).row
    
    Dim dictClients As Dictionary 'Code, Nom du Client
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowClient
        dictClients.add CStr(referenceSheet.Cells(i, 2).value), referenceSheet.Cells(i, 1).value
    Next i
    
    Dim codeClient As String, nomClient As String, updatedNomClient As String
    For i = 2 To lastUsedRow
        codeClient = sourceSheet.Cells(i, 5).value
        nomClient = sourceSheet.Cells(i, 6).value
        updatedNomClient = dictClients(codeClient)
        Debug.Print i & " : " & codeClient & " - " & nomClient & " ---> " & updatedNomClient
        sourceSheet.Cells(i, 6).value = updatedNomClient
    Next i
    
    'Save and close the destination workbook
    sourceWorkbook.Save
    sourceWorkbook.Close
    
    'Clean up
    Set sourceSheet = Nothing
    Set sourceRange = Nothing
    Set sourceWorkbook = Nothing
    Set referenceSheet = Nothing
    Set referenceWorkbook = Nothing
    
    MsgBox "Les donn�es ont �t� copi�es avec succ�s dans le fichier destination."
    
End Sub

Sub Adjust_Client_Name_In_CAR()  '2024-08-07 @ 17:11

    Dim sourceRange As Range
    
    'D�finir les chemins d'acc�s des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER.xlsx"
    Dim clientMF As String
    clientMF = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("CAR")
    
    'D�termine la derni�re rang�e utilis�e dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.rows.count, 1).End(xlUp).row
    Dim lastUsedCol As Long
    lastUsedCol = sourceSheet.Cells(1, sourceSheet.columns.count).End(xlToLeft).Column
    
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
        dictClients.add CStr(referenceSheet.Cells(i, 2).value), referenceSheet.Cells(i, 1).value
'        Debug.Print referenceSheet.Cells(i, 2).value & " - " & referenceSheet.Cells(i, 1).value
    Next i
    
    Dim codeClient As String, nomClient As String, updatedNomClient As String
    For i = 3 To lastUsedRow
        codeClient = sourceSheet.Cells(i, 4).value
        nomClient = sourceSheet.Cells(i, 3).value
        updatedNomClient = dictClients(codeClient)
        Debug.Print i & " : " & codeClient & " - " & nomClient & " ---> " & updatedNomClient
        sourceSheet.Cells(i, 3).value = updatedNomClient
    Next i
    
    'Save and close the destination workbook
    sourceWorkbook.Save
    sourceWorkbook.Close
    
    'Clean up
    Set sourceSheet = Nothing
    Set sourceRange = Nothing
    Set sourceWorkbook = Nothing
    Set referenceSheet = Nothing
    Set referenceWorkbook = Nothing
    
    MsgBox "Les donn�es ont �t� copi�es avec succ�s dans le fichier destination."
    
End Sub

Sub Check_Client_Name() '2024-08-10 @ 10:13

    'D�finir les chemins d'acc�s des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("Clients")
    
    'D�termine la derni�re rang�e utilis�e dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.rows.count, 1).End(xlUp).row
    
    Dim codeClient As String, nomClient As String, contactFact As String
    Dim i As Long
    For i = 2 To lastUsedRow
        codeClient = sourceSheet.Cells(i, 2).value
        nomClient = Trim(sourceSheet.Cells(i, 1).value)
        contactFact = Trim(sourceSheet.Cells(i, 3).value)
        If InStr(nomClient, contactFact) = 0 Then
            Debug.Print i & " : " & codeClient & " - " & nomClient & " on ajoute '" & contactFact & "'"
        End If
    Next i
    
    'Save and close the destination workbook
    sourceWorkbook.Save
    sourceWorkbook.Close
    
    'Clean up
    Set sourceSheet = Nothing
    Set sourceWorkbook = Nothing
    
    MsgBox "Les donn�es ont �t� v�rifi�es avec succ�s dans le fichier Clients."
    
End Sub

Sub Temp_Build_Hours_Summary() '2024-08-12 @ 21:09

    'D�finir les chemins d'acc�s des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\GCF_DataFiles\GCF_BD_MASTER.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("TEC_Local")
    
    'D�termine la derni�re rang�e utilis�e dans le fichier Source
    Dim lastUsedRow As Long
    lastUsedRow = sourceSheet.Cells(sourceSheet.rows.count, 1).End(xlUp).row
    
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
    
    'Clean up
    Set sourceSheet = Nothing
    Set sourceWorkbook = Nothing
    
    MsgBox "Sommaire des heures est compl�t�."
    
End Sub

Sub Fix_Client_Name_In_TEC()  '2024-08-23 @ 06:32

    'Source - D�finir les chemins d'acc�s des fichiers, le Workbook, le Worksheet et le Range
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Master.xlsx"
    Dim wbSource As Workbook: Set wbSource = Workbooks.Open(sourceFilePath)
    Dim wsSource As Worksheet: Set wsSource = wbSource.Worksheets("TEC_Local")
    
    'D�termine la derni�re rang�e et derni�re colonne utilis�es dans wshTEC_Local
    Dim lastUsedRowTEC As Long
    lastUsedRowTEC = wsSource.Cells(wsSource.rows.count, 1).End(xlUp).row
    
    'Open the Master File Workbook
    Dim clientMFPath As String
    clientMFPath = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    Dim wbMF As Workbook: Set wbMF = Workbooks.Open(clientMFPath)
    Dim wsMF As Worksheet: Set wsMF = wbMF.Worksheets("Clients")
    Dim lastUsedRowTECClient As Long
    lastUsedRowTECClient = wsMF.Cells(wsMF.rows.count, "A").End(xlUp).row
    
    'Setup output file
    Dim strOutput As String
    strOutput = "TEC_Correction_Nom"
    Call CreateOrReplaceWorksheet(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").value = "TEC_Nom_Client"
    wsOutput.Range("B1").value = "Code_de_Client"
    wsOutput.Range("C1").value = "Nom_Client_Master"
    wsOutput.Range("D1").value = "TEC_ID"
    wsOutput.Range("E1").value = "TEC_Prof"
    wsOutput.Range("F1").value = "TEC_Date"
    Call Make_It_As_Header(wsOutput.Range("A1:F1"))
    
    'Build the dictionnary (Code, Nom du client) from Client's Master File
    Dim dictClients As Dictionary
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowTECClient
        dictClients.add CStr(wsMF.Cells(i, 2).value), wsMF.Cells(i, 1).value
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
            Debug.Print i & " : " & codeClientTEC & " - " & nomClientTEC & " <---> " & nomClientFromMF
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
    
    wsOutput.columns.AutoFit

    'Result print setup - 2024-08-05 @ 05:16
    rowOutput = rowOutput + 1
    wsOutput.Range("A" & rowOutput).value = "**** " & Format$(lastUsedRowTEC - 1, "###,##0") & _
                                        " lignes analys�es dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("A2:F" & rowOutput)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:E" & rowOutput)
    Dim header1 As String: header1 = "Correction des noms de clients dans les TEC"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "P")
    
    'Close the 2 workbooks without saving anything
    wbSource.Close saveChanges:=True
    wbMF.Close saveChanges:=False

    'Clean up
    Set wsSource = Nothing
    Set wbSource = Nothing
    Set wsMF = Nothing
    Set wbMF = Nothing
    
    MsgBox "Il y a " & casDelta & " cas o� le nom du client (TEC) diff�re" & _
            vbNewLine & vbNewLine & "du nom de client du Fichier MA�TRE", vbInformation
    
End Sub

