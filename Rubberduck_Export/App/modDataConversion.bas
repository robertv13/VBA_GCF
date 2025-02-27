Attribute VB_Name = "modDataConversion"
Option Explicit

'Importation des clients � partir de ... \DataConversion\Clients.xlsx
Sub Copy_Data_Between_Closed_Workbooks_Clients() '2024-08-03 @ 09:40

    Stop 'One shot deal !!!
    
    Dim sourceRange As Range
    
    'D�finir les chemins d'acc�s des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\DataConversion\Clients.xlsx"
    Dim destinationFilePath As String
    destinationFilePath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & "GCF_BD_Entr�e.xlsx"
    
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
    sourceWorkbook.Close SaveChanges:=False
    
    'Lib�rer la m�moire
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
    Dim wb As Workbook: Set wb = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx")

    'Define the worksheet you want to work with
    Dim ws As Worksheet: Set ws = wb.Worksheets("Clients")
    
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
        client = ws.Cells(i, 1).Value
        client_ID = ws.Cells(i, 2).Value
        contactFacturation = ws.Cells(i, 3).Value
        
        'Process the data and make adjustments if necessary
        posOpenParenthesis = InStr(client, "(")
        posCloseParenthesis = InStr(client, ")")
        numberOpenParenthesis = CountCharOccurrences(client, "(")
        numberCloseParenthesis = CountCharOccurrences(client, ")")
        
        If numberOpenParenthesis = 1 And numberCloseParenthesis = 1 Then
            If posCloseParenthesis > posOpenParenthesis + 5 Then
                client = Replace(client, "(", "[")
                client = Replace(client, ")", "]")
                ws.Cells(i, 1).Value = client
                Debug.Print i & " - " & client
            End If
        End If
        
    Next i
    
    wb.Save
    
    'Lib�rer la m�moire
    Set wb = Nothing
    Set ws = Nothing
    
    MsgBox "Le traitement est compl�t� sur " & i - 1 & " lignes"
    
End Sub

'Ajustements � la feuille DB_Clients (Ajout du contactdans le nom du client)
Sub Clients_Ajout_Contact_Dans_Nom()

    'Declare and open the closed workbook
    Dim wb As Workbook: Set wb = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx")

    'Define the worksheet you want to work with
    Dim ws As Worksheet: Set ws = wb.Worksheets("Clients")
    
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
        client = ws.Cells(i, fClntMFClientNom).Value
        client_ID = ws.Cells(i, fClntMFClient_ID).Value
        contactFacturation = Trim(ws.Cells(i, fClntMFContactFacturation).Value)
        
        'Process the data and make adjustments if necessary
        posOpenSquareBracket = InStr(client, "[")
        posCloseSquareBracket = InStr(client, "]")
        
        If numberOpenSquareBracket = 0 And numberCloseSquareBracket = 0 Then
            If contactFacturation <> "" And InStr(client, contactFacturation) = 0 Then
                client = Trim(client) & " [" & contactFacturation & "]"
                ws.Cells(i, 1).Value = client
                Debug.Print i & " - " & client
            End If
        End If
        
    Next i
    
    wb.Save
    
    'Lib�rer la m�moire
    Set wb = Nothing
    Set ws = Nothing
    
    MsgBox "Le traitement est compl�t� sur " & i - 1 & " lignes"
    
End Sub

Sub Import_Data_From_Closed_Workbooks_TEC() '2024-08-14 @ 06:43 & 2024-08-03 @ 16:15

    Stop 'One shot deal !!!
    
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
        prof = Trim(rst.Fields(0).Value)
        clientCode = Trim(rst.Fields(2).Value)
'        clientCode = Left(client, 10)
'            clientCode = Left(clientCode, InStr(clientCode, " -") - 1)
        client = Trim(rst.Fields(3).Value)
'        client = Mid(client, InStr(client, " - ") + 3, Len(client))
        totHres = totHres + CDbl(rst.Fields(5).Value)
        
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
        wsDest.Range("A" & rowNum).Value = TEC_ID
        wsDest.Range("B" & rowNum).Value = Get_ID_From_Initials(prof)
        wsDest.Range("C" & rowNum).Value = prof
        wsDest.Range("D" & rowNum).Value = rst.Fields(1).Value
        wsDest.Range("E" & rowNum).Value = clientCode
        wsDest.Range("F" & rowNum).Value = client
        wsDest.Range("G" & rowNum).Value = rst.Fields(4).Value
        wsDest.Range("H" & rowNum).Value = rst.Fields(5).Value
        wsDest.Range("I" & rowNum).Value = ""
        wsDest.Range("J" & rowNum).Value = "VRAI"
        wsDest.Range("K" & rowNum).Value = Format$(Now(), "dd/mm/yyyy hh:nn:ss")
        wsDest.Range("L" & rowNum).Value = "FAUX"
        wsDest.Range("M" & rowNum).Value = ""
        wsDest.Range("N" & rowNum).Value = "FAUX"
        wsDest.Range("O" & rowNum).Value = ThisWorkbook.Name
        wsDest.Range("P" & rowNum).Value = ""
        
        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les TEC ont �t� import�s, pour un total de " & totHres & " heures"
    End If
    
    'Lib�rer la m�moire
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Set rng = Nothing
    Set wsDest = Nothing
    
End Sub

'Only valid for this conversion process
Function Get_ID_From_Initials(p As String) As Long

    Stop 'One shot deal
    
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

    Stop 'One shot deal
    
    'D�finir les chemins d'acc�s des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\DataConversion\Fournisseurs.xlsx"
    Dim destinationFilePath As String
    destinationFilePath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & "GCF_BD_Entr�e.xlsx"
    
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
    sourceWorkbook.Close SaveChanges:=False
    
    'Lib�rer la m�moire
    Set sourceSheet = Nothing
    Set sourceRange = Nothing
    Set sourceWorkbook = Nothing
    Set destinationSheet = Nothing
    Set destinationWorkbook = Nothing
    
    MsgBox "Les donn�es (fournisseurs) ont �t� copi�es avec succ�s dans" & vbNewLine & _
            vbNewLine & "le fichier destination."
    
End Sub

Sub Import_Data_From_Closed_Workbooks_GL_BV() '2024-08-03 @ 18:20

    Stop 'One shot deal
    
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
        descriptionGL = rst.Fields(0).Value
        codeGL = Fn_GetGL_Code_From_GL_Description(descriptionGL)
        amount = rst.Fields(1).Value
        If amount > 0 Then
            totalDT = totalDT + amount
        Else
            totalCT = totalCT - amount
        End If
        
        wsDest.Range("A" & rowNum).Value = No_Entr�e
        wsDest.Range("B" & rowNum).Value = "07/31/2024"
        wsDest.Range("C" & rowNum).Value = "Solde de fermeture (conversion)"
        wsDest.Range("D" & rowNum).Value = "Conv."
        wsDest.Range("E" & rowNum).Value = codeGL
        wsDest.Range("F" & rowNum).Value = descriptionGL
        If amount >= 0 Then
            wsDest.Range("G" & rowNum).Value = amount
        Else
            wsDest.Range("H" & rowNum).Value = -amount
        End If
        wsDest.Range("I" & rowNum).Value = ""
        wsDest.Range("J" & rowNum).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        
        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les soldes ont �t� import�s, pour un total d�bit de " & totalDT & _
                vbNewLine & vbNewLine & "un total cr�dit de " & totalCT
    End If
    
    'Lib�rer la m�moire
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Set wsDest = Nothing
    
End Sub

Sub Import_Data_From_Closed_Workbooks_CC() '2024-08-04 @ 07:31

    Stop 'One shot deal
    
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
        client = rst.Fields(0).Value
        dateFact = rst.Fields(1).Value
        factNo = rst.Fields(2).Value
        totalFact = rst.Fields(3).Value
        recu = rst.Fields(4).Value
        If IsNull(rst.Fields(5).Value) Then
            dateRecu = ""
        Else
            dateRecu = rst.Fields(5).Value
        End If
        solde = rst.Fields(6).Value
        
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
        
        wsDest.Range("A" & rowNum).Value = factNo
        wsDest.Range("B" & rowNum).Value = dateFact
        wsDest.Range("C" & rowNum).Value = client
        wsDest.Range("D" & rowNum).Value = clientCode
        wsDest.Range("E" & rowNum).Value = "Unpaid"
        wsDest.Range("F" & rowNum).Value = "Net 30"
        dateDue = DateAdd("d", 30, CDate(dateFact))
        wsDest.Range("G" & rowNum).Value = Format$(CDate(dateDue), "mm/dd/yyyy")
        wsDest.Range("H" & rowNum).Value = totalFact
        wsDest.Range("I" & rowNum).Value = recu
        wsDest.Range("J" & rowNum).Value = totalFact - recu
        joursDue = DateDiff("d", dateDue, Date)
        wsDest.Range("K" & rowNum).Value = joursDue
        rowNum = rowNum + 1

        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les CAR ont �t� import�s, pour un total de " & Format$(totCAR, "#,##0.00$")
    End If
    
    'Lib�rer la m�moire
    rst.Close
    cnn.Close
    Set rng = Nothing
    Set rst = Nothing
    Set cnn = Nothing
    Set wsDest = Nothing
    
End Sub

Sub Compare_2_Excel_Files() '------------------------------------------ 2024-09-02 @ 06:24
    
    Application.ScreenUpdating = False
    
    'Declare and open the 2 workbooks
    Dim wbWas As Workbook
    Set wbWas = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx", ReadOnly:=True)
    Debug.Print wbWas.Name
    Dim wbNow As Workbook
    Set wbNow = Workbooks.Open("C:\VBA\GC_FISCALIT�\GCF_DataFiles\2024_09_01_1835\GCF_BD_Entr�e_TBA.xlsx", ReadOnly:=True)
    Debug.Print wbNow.Name

    'Declare the 2 worksheets
    Dim wsWas As Worksheet
    Set wsWas = wbWas.Worksheets("Clients")
    Dim wsNow As Worksheet
    Set wsNow = wbNow.Worksheets("Clients")
    
    'D�termine la derni�re ligne utilis�e dans chacune des 2 feuilles
    Dim lastUsedRowWas As Long
    lastUsedRowWas = wsWas.Cells(wsWas.rows.count, 1).End(xlUp).row
    Dim lastUsedRowNOw As Long
    lastUsedRowNOw = wsNow.Cells(wsNow.rows.count, 1).End(xlUp).row
    
    'D�termine le nombre de colonnes dans l'ancienne feuille
    Dim lastUsedColWas As Long
    lastUsedColWas = wsWas.Cells(wsWas.columns.count).End(xlToLeft).Column
    
    'Erase and create a new worksheet for differences
    Dim wsNameStr As String
    wsNameStr = "X_Diff�rences"
    Dim wsDiff As Worksheet
    Call CreateOrReplaceWorksheet(wsNameStr)
    Set wsDiff = ThisWorkbook.Worksheets(wsNameStr)
    wsDiff.Range("A1").Value = "Ligne"
    wsDiff.Range("B1").Value = "Colonne"
    wsDiff.Range("C1").Value = "CodeClient"
    wsDiff.Range("D1").Value = "Nom du Client"
    wsDiff.Range("E1").Value = "Avant changement"
    wsDiff.Range("F1").Value = "Type"
    wsDiff.Range("G1").Value = "Apr�s changement"
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
        clientCode = CStr(wsWas.Cells(i, 2).Value)
        'Trouver la ligne correspondante dans la nouvelle version
        Set foundRow = wsNow.columns(2).Find(What:=clientCode, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundRow Is Nothing Then
            Debug.Print "Ligne : " & i
            'Comparer les cellules des lignes correspondantes
            For j = 1 To lastUsedColWas
                readCells = readCells + 1
                Set cellWas = wsWas.Cells(i, j)
                Set cellNow = wsNow.Cells(foundRow.row, j)
                If CStr(cellWas.Value) <> CStr(cellNow.Value) Then
                    wsDiff.Cells(diffRow, 1).Value = i
                    wsDiff.Cells(diffRow, 2).Value = j
                    wsDiff.Cells(diffRow, 3).Value = wsWas.Cells(i, 2).Value
                    wsDiff.Cells(diffRow, 4).Value = wsWas.Cells(cellWas.row, 1).Value
                    wsDiff.Cells(diffRow, 5).Value = cellWas.Value
                    wsDiff.Cells(diffRow, 6).Value = "'--->"
                    wsDiff.Cells(diffRow, 7).Value = cellNow.Value
                    diffRow = diffRow + 1
                End If
            Next j
        Else
            wsDiff.Cells(diffRow, 1).Value = i
            wsDiff.Cells(diffRow, 3).Value = wsWas.Cells(i, 2).Value
            wsDiff.Cells(diffRow, 4).Value = wsWas.Cells(cellWas.row, 1).Value
            wsDiff.Cells(diffRow, 5).Value = cellWas.Value
            wsDiff.Cells(diffRow, 6).Value = "XXXX"
            diffRow = diffRow + 1
        End If
    Next i
            
    wsDiff.columns.AutoFit
    
    'Result print setup - 2024-08-05 @ 05:16
    diffRow = diffRow + 1
    wsDiff.Range("A" & diffRow).Value = "**** " & Format$(readCells, "###,##0") & _
                                        " cellules analys�es dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:G" & diffRow)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:DC" & diffRow)
    Dim header1 As String: header1 = "V�rification des diff�rences"
    Dim header2 As String: header2 = "Clients"
    Call Simple_Print_Setup(wsDiff, rngToPrint, header1, header2, "$1:$1", "P")
    
    Application.ScreenUpdating = True
    
    wsDiff.Activate

    'Close the workbooks without saving
    wbWas.Close SaveChanges:=False
    wbNow.Close SaveChanges:=False
    
    'Lib�rer la m�moire
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
    
    MsgBox "La comparaison est compl�t�e." & vbNewLine & vbNewLine & _
           differences, vbInformation
           
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
        dictClients.Add CStr(referenceSheet.Cells(i, 2).Value), referenceSheet.Cells(i, 1).Value
    Next i
    
    Dim codeClient As String, nomClient As String, updatedNomClient As String
    For i = 2 To lastUsedRow
        codeClient = sourceSheet.Cells(i, 5).Value
        nomClient = sourceSheet.Cells(i, 6).Value
        updatedNomClient = dictClients(codeClient)
        Debug.Print i & " : " & codeClient & " - " & nomClient & " ---> " & updatedNomClient
        sourceSheet.Cells(i, 6).Value = updatedNomClient
    Next i
    
    'Save and close the destination workbook
    sourceWorkbook.Save
    sourceWorkbook.Close
    
    'Lib�rer la m�moire
    Set dictClients = Nothing
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
        dictClients.Add CStr(referenceSheet.Cells(i, 2).Value), referenceSheet.Cells(i, 1).Value
'        Debug.Print referenceSheet.Cells(i, 2).Value & " - " & referenceSheet.Cells(i, 1).Value
    Next i
    
    Dim codeClient As String, nomClient As String, updatedNomClient As String
    For i = 3 To lastUsedRow
        codeClient = sourceSheet.Cells(i, 4).Value
        nomClient = sourceSheet.Cells(i, 3).Value
        updatedNomClient = dictClients(codeClient)
        Debug.Print i & " : " & codeClient & " - " & nomClient & " ---> " & updatedNomClient
        sourceSheet.Cells(i, 3).Value = updatedNomClient
    Next i
    
    'Save and close the destination workbook
    sourceWorkbook.Save
    sourceWorkbook.Close
    
    'Lib�rer la m�moire
    Set dictClients = Nothing
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
        codeClient = sourceSheet.Cells(i, fClntMFClient_ID).Value
        nomClient = Trim(sourceSheet.Cells(i, fClntMFClientNom).Value)
        contactFact = Trim(sourceSheet.Cells(i, fClntMFContactFacturation).Value)
        If InStr(nomClient, contactFact) = 0 Then
            Debug.Print i & " : " & codeClient & " - " & nomClient & " on ajoute '" & contactFact & "'"
        End If
    Next i
    
    'Save and close the destination workbook
    sourceWorkbook.Save
    sourceWorkbook.Close
    
    'Lib�rer la m�moire
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
        profID = sourceSheet.Cells(i, 2).Value
        prof = sourceSheet.Cells(i, 3).Value
        dateServ = sourceSheet.Cells(i, 4).Value
        codeClient = sourceSheet.Cells(i, 5).Value
        nomClient = Trim(sourceSheet.Cells(i, 6).Value)
        hresSaisies = Trim(sourceSheet.Cells(i, 8).Value)
        estFacturable = sourceSheet.Cells(i, 10).Value
        estFacturee = sourceSheet.Cells(i, 12).Value
        estDetruit = sourceSheet.Cells(i, 14).Value
        
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
    
    'Lib�rer la m�moire
    Set sourceSheet = Nothing
    Set sourceWorkbook = Nothing
    
    MsgBox "Sommaire des heures est compl�t�."
    
End Sub

Sub Fix_Client_Name_In_TEC()  '2024-08-23 @ 06:32

    'Source - D�finir les chemins d'acc�s des fichiers, le Workbook, le Worksheet et le Range
    Dim sourceFilePath As String
    sourceFilePath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Master.xlsx"
    Dim wbSource As Workbook: Set wbSource = Workbooks.Open(sourceFilePath)
    Dim wsSource As Worksheet: Set wsSource = wbSource.Worksheets("TEC_Local")
    
    'D�termine la derni�re rang�e et derni�re colonne utilis�es dans wshTEC_Local
    Dim lastUsedRowTEC As Long
    lastUsedRowTEC = wsSource.Cells(wsSource.rows.count, 1).End(xlUp).row
    
    'Open the Master File Workbook
    Dim clientMFPath As String
    clientMFPath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx"
    Dim wbMF As Workbook: Set wbMF = Workbooks.Open(clientMFPath)
    Dim wsMF As Worksheet: Set wsMF = wbMF.Worksheets("Clients")
    Dim lastUsedRowTECClient As Long
    lastUsedRowTECClient = wsMF.Cells(wsMF.rows.count, "A").End(xlUp).row
    
    'Setup output file
    Dim strOutput As String
    strOutput = "X_TEC_Correction_Nom"
    Call CreateOrReplaceWorksheet(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").Value = "TEC_Nom_Client"
    wsOutput.Range("B1").Value = "Code_de_Client"
    wsOutput.Range("C1").Value = "Nom_Client_Master"
    wsOutput.Range("D1").Value = "TEC_ID"
    wsOutput.Range("E1").Value = "TEC_Prof"
    wsOutput.Range("F1").Value = "TEC_Date"
    Call Make_It_As_Header(wsOutput.Range("A1:F1"))
    
    'Build the dictionnary (Code, Nom du client) from Client's Master File
    Dim dictClients As Dictionary
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowTECClient
        dictClients.Add CStr(wsMF.Cells(i, 2).Value), wsMF.Cells(i, 1).Value
    Next i
    
    'Parse TEC_Local to verify TEC's clientName vs. MasterFile's clientName
    Dim codeClientTEC As String, nomClientTEC As String, nomClientFromMF As String
    Dim casDelta As Long, rowOutput As Long
    rowOutput = 2
    For i = 2 To lastUsedRowTEC
        codeClientTEC = wsSource.Cells(i, 5).Value
        nomClientTEC = wsSource.Cells(i, 6).Value
        nomClientFromMF = dictClients(codeClientTEC)
        If nomClientTEC <> nomClientFromMF Then
            Debug.Print i & " : " & codeClientTEC & " - " & nomClientTEC & " <---> " & nomClientFromMF
            wsSource.Cells(i, 6).Value = nomClientFromMF
            wsOutput.Cells(rowOutput, 1).Value = nomClientTEC
            wsOutput.Cells(rowOutput, 2).Value = codeClientTEC
            wsOutput.Cells(rowOutput, 3).Value = nomClientFromMF
            wsOutput.Cells(rowOutput, 4).Value = wsSource.Cells(i, 1).Value
            wsOutput.Cells(rowOutput, 5).Value = wsSource.Cells(i, 3).Value
            wsOutput.Cells(rowOutput, 6).Value = wsSource.Cells(i, 4).Value
            rowOutput = rowOutput + 1
            casDelta = casDelta + 1
        End If
    Next i
    
    wsOutput.columns.AutoFit

    'Result print setup - 2024-08-05 @ 05:16
    rowOutput = rowOutput + 1
    wsOutput.Range("A" & rowOutput).Value = "**** " & Format$(lastUsedRowTEC - 1, "###,##0") & _
                                        " lignes analys�es dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("A2:F" & rowOutput)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:E" & rowOutput)
    Dim header1 As String: header1 = "Correction des noms de clients dans les TEC"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wbSource.Close SaveChanges:=True
    wbMF.Close SaveChanges:=False

    'Lib�rer la m�moire
    Set dictClients = Nothing
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set wsSource = Nothing
    Set wsOutput = Nothing
    Set wbSource = Nothing
    Set wsMF = Nothing
    Set wbMF = Nothing
    
    MsgBox "Il y a " & casDelta & " cas o� le nom du client (TEC) diff�re" & _
            vbNewLine & vbNewLine & "du nom de client du Fichier MA�TRE", vbInformation
    
End Sub

Public Sub Fix_Client_Name_In_CAR()  '2024-08-31 @ 06:52

    'Worksheets to be corrected - Open the workbook (worksheet will be determined later)
    Dim sourceFilePath As String
    sourceFilePath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Master.xlsx"
    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(sourceFilePath)
    
    'Client's Master File - Workbook & Worksheet
    Dim clientMFPath As String
    clientMFPath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx"
    Dim wbMF As Workbook
    Set wbMF = Workbooks.Open(clientMFPath)
    Dim wsMF As Worksheet
    Set wsMF = wbMF.Worksheets("Clients")
    Dim lastUsedRowClientMF As Long
    lastUsedRowClientMF = wsMF.Cells(wsMF.rows.count, "A").End(xlUp).row
    
    'Build the dictionnary (Code, Nom du client) from Client's Master File
    Dim clientName As String
    Dim dictClients As Dictionary
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowClientMF
        'Enl�ve les informations de contact
        clientName = wsMF.Cells(i, 1).Value
        Do While InStr(clientName, "[") > 0 And InStr(clientName, "]") > 0
            clientName = Fn_Strip_Contact_From_Client_Name(clientName)
        Loop
        dictClients.Add CStr(wsMF.Cells(i, 2).Value), clientName
    Next i
    
    'Setup output file
    Dim strOutput As String
    strOutput = "X_CAR_Correction_Nom"
    Call CreateOrReplaceWorksheet(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").Value = "Feuille"
    wsOutput.Range("B1").Value = "No Facture"
    wsOutput.Range("C1").Value = "Nom de client (Facture)"
    wsOutput.Range("D1").Value = "Code_de_Client"
    wsOutput.Range("E1").Value = "Chang� pour"
    Call Make_It_As_Header(wsOutput.Range("A1:E1"))
    Dim rowOutput As Long
    rowOutput = 2 'Skip the header
    
    'There is 2 worksheets to adjust
    Dim ws As Variant
    Dim wsSource As Worksheet
    Dim casDelta As Long
    For Each ws In Array("FAC_Ent�te|4|6", "FAC_Comptes_Clients|4|3")
        Dim param As String
        param = Mid(ws, InStr(ws, "|") + 1)
        ws = Left(ws, InStr(ws, "|") - 1)
        Dim colClientID As Long, colClientName As Long
        colClientID = Left(param, InStr(param, "|") - 1)
        colClientName = Right(param, Len(param) - InStr(param, "|"))
        'Set the worksheet Object
        Set wsSource = wbSource.Sheets(ws)
        'D�termine la derni�re rang�e utilis�e dans la feuille
        Dim lastUsedRow As Long
        lastUsedRow = wsSource.Cells(wsSource.rows.count, 1).End(xlUp).row
        Dim codeClientCAR As String, nomClientCAR As String, nomClientFromMF As String
        For i = 3 To lastUsedRow
            codeClientCAR = wsSource.Cells(i, colClientID).Value
            nomClientCAR = wsSource.Cells(i, colClientName).Value
            nomClientFromMF = dictClients(codeClientCAR)
            If nomClientCAR <> nomClientFromMF Then
                Debug.Print i & " : " & codeClientCAR & " - " & nomClientCAR & " <---> " & nomClientFromMF
                wsSource.Cells(i, colClientName).Value = nomClientFromMF
                wsOutput.Cells(rowOutput, 1).Value = wsSource.Name
                wsOutput.Cells(rowOutput, 2).Value = wsSource.Cells(i, 1).Value
                wsOutput.Cells(rowOutput, 3).Value = nomClientCAR
                wsOutput.Cells(rowOutput, 4).Value = codeClientCAR
                wsOutput.Cells(rowOutput, 5).Value = nomClientFromMF
                rowOutput = rowOutput + 1
                casDelta = casDelta + 1
            End If
        Next i
    Next ws
    
    wsOutput.columns.AutoFit

    'Result print setup - 2024-08-05 @ 05:16
    rowOutput = rowOutput + 1
    wsOutput.Range("A" & rowOutput).Value = "**** " & Format$(lastUsedRow - 1, "###,##0") & _
                                        " lignes analys�es dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("A2:F" & rowOutput)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:E" & rowOutput)
    Dim header1 As String: header1 = "Correction des noms de clients dans les CAR"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wbSource.Close SaveChanges:=True
    wbMF.Close SaveChanges:=False

    'Lib�rer la m�moire
    Set dictClients = Nothing
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set wsOutput = Nothing
    Set wsSource = Nothing
    Set wbSource = Nothing
    Set ws = Nothing
    Set wsMF = Nothing
    Set wbMF = Nothing
    
''    MsgBox "Il y a " & casDelta & " cas o� le nom du client (TEC) diff�re" & _
''            vbNewLine & vbNewLine & "du nom de client du Fichier MA�TRE", vbInformation
'
End Sub

Sub Import_Missing_AR_Records() '2024-08-24 @ 15:58

    Application.ScreenUpdating = False
    
    'Declare and open the 2 workbooks
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER - Copie.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER.xlsx")

    'Declare the 2 worksheets
    Dim ws1 As Worksheet: Set ws1 = wb1.Worksheets("FAC_Comptes_Clients")
    Dim ws2 As Worksheet: Set ws2 = wb2.Worksheets("FAC_Ent�te")
    
    Dim lastUsedRow As Long
    lastUsedRow = ws1.Cells(ws1.rows.count, "A").End(xlUp).row
    Dim row As Long, rowFAC_Entete As Long
    row = 2
    rowFAC_Entete = ws2.Cells(ws2.rows.count, "A").End(xlUp).row + 1
    
    Dim i As Integer
    For i = 2 To lastUsedRow
        If InStr(ws1.Range("A" & i).Value, "24-") <> 1 Then
            ws2.Range("A" & rowFAC_Entete).Value = ws1.Range("A" & i)
            ws2.Range("B" & rowFAC_Entete).Value = ws1.Range("B" & i)
            ws2.Range("C" & rowFAC_Entete).Value = "C"
            ws2.Range("D" & rowFAC_Entete).Value = ws1.Range("D" & i)
            ws2.Range("F" & rowFAC_Entete).Value = ws1.Range("C" & i)
            
            ws2.Range("J" & rowFAC_Entete).Value = ws1.Range("H" & i)
            ws2.Range("K" & rowFAC_Entete).Value = "Frais de poste"
            ws2.Range("L" & rowFAC_Entete).Value = 0
            ws2.Range("M" & rowFAC_Entete).Value = "Frais d'expert en taxes"
            ws2.Range("N" & rowFAC_Entete).Value = 0
            ws2.Range("O" & rowFAC_Entete).Value = "Autres frais"
            ws2.Range("P" & rowFAC_Entete).Value = 0
            
            ws2.Range("Q" & rowFAC_Entete).Value = Format$(CDbl(5), "#0.000")
            ws2.Range("R" & rowFAC_Entete).Value = 0
            ws2.Range("S" & rowFAC_Entete).Value = Format$(CDbl(9.975), "#0.000")
            ws2.Range("T" & rowFAC_Entete).Value = 0
            
            ws2.Range("U" & rowFAC_Entete).Value = ws1.Range("H" & i)
            ws2.Range("V" & rowFAC_Entete).Value = 0
            rowFAC_Entete = rowFAC_Entete + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    'Close the workbooks without saving
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=True
    
    'Lib�rer la m�moire
    Set wb1 = Nothing
    Set wb2 = Nothing
    Set ws1 = Nothing
    Set ws2 = Nothing
    
    MsgBox "Le traitement est compl�t�", vbInformation
           
End Sub

Sub Merge_Missing_AR_Records() '2024-08-29 @ 07:29

    Application.ScreenUpdating = False
    
    'Declare and open the 2 workbooks
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALIT�\GCF_DataFiles\CAR_A_COMPL�TER.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER.xlsx")

    'Declare the 2 worksheets
    Dim ws1 As Worksheet: Set ws1 = wb1.Worksheets("Feuil1")
    Dim ws2 As Worksheet: Set ws2 = wb2.Worksheets("FAC_Ent�te")
    
    Dim lastUsedRow As Long
    lastUsedRow = ws1.Cells(ws1.rows.count, "A").End(xlUp).row
    Dim lastUsedRowTarget As Long
    lastUsedRowTarget = ws2.Cells(ws2.rows.count, "A").End(xlUp).row
    
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
        invNo = ws1.Cells(i, 1).Value
        
        hono = ws1.Cells(i, 10).Value
        af1Str = ws1.Cells(i, 11).Value
        af1 = ws1.Cells(i, 12).Value
        af2Str = ws1.Cells(i, 13).Value
        af2 = ws1.Cells(i, 14).Value
        af3Str = ws1.Cells(i, 15).Value
        af3 = ws1.Cells(i, 16).Value
        tps = ws1.Cells(i, 18).Value
        tvq = ws1.Cells(i, 20).Value
        arTotal = ws1.Cells(i, 21).Value
        depot = ws1.Cells(i, 22).Value

        If hono + af1 + af2 + af3 + tps + tvq <> arTotal Then
            MsgBox "Une ligne (" & i & ") ne balance pas !!!", vbCritical
        End If
        
        'Find the InvNo in wshFAC_Ent�te
        Set foundCells = rngTarget.columns(1).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
        If foundCells Is Nothing Then
            MsgBox "**** Je n'ai pas trouv� la facture '" & invNo & "' dans wshFAC_Ent�te", vbCritical
        Else
            ii = foundCells.row
        End If
        
        If ws2.Cells(ii, 21).Value <> arTotal Then
            MsgBox "Probl�me d'int�grit� pour la facture '" & invNo & "' au niveau de arTotal", vbCritical
        End If
        
        'Replace values in Target, with the Source info
        ws2.Cells(ii, 10).Value = hono
        If af1 <> 0 And af1Str <> ws2.Cells(ii, 11) Then
            ws2.Cells(ii, 11).Value = af1Str
        End If
        ws2.Cells(ii, 12).Value = af1
        
        If af2 <> 0 And af2Str <> ws2.Cells(ii, 13) Then
            ws2.Cells(ii, 13).Value = af2Str
        End If
        ws2.Cells(ii, 14).Value = af2
        
        If af3 <> 0 And af3Str <> ws2.Cells(ii, 15) Then
            ws2.Cells(ii, 15).Value = af3Str
        End If
        ws2.Cells(ii, 16).Value = af3
        
        ws2.Cells(ii, 18).Value = tps
        ws2.Cells(ii, 20).Value = tvq
        
        ws2.Cells(ii, 22).Value = depot
        
        If ws2.Cells(ii, 10) + ws2.Cells(ii, 12) + ws2.Cells(ii, 14) + ws2.Cells(ii, 16) + _
            ws2.Cells(ii, 18) + ws2.Cells(ii, 20) <> ws2.Cells(ii, 21).Value Then
            MsgBox "Probl�me avec les assignations...", vbCritical
        End If
        
        t(1) = t(1) + ws2.Cells(ii, 10)
        t(2) = t(2) + ws2.Cells(ii, 12)
        t(3) = t(3) + ws2.Cells(ii, 14)
        t(4) = t(4) + ws2.Cells(ii, 16)
        t(5) = t(5) + ws2.Cells(ii, 18)
        t(6) = t(6) + ws2.Cells(ii, 20)
        t(7) = t(7) + ws2.Cells(ii, 21)
        t(8) = t(8) + ws2.Cells(ii, 22)
        
        Debug.Print "x8", invNo, ii, Format(i / lastUsedRow, "##0.00 %")
    
    Next i
    
    Debug.Print t(1), t(2), t(3), t(4), t(5), t(6), t(7), t(8)
    
    Application.ScreenUpdating = True
    
    'Close the workbooks without saving
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=True
    
    'Lib�rer la m�moire
    Set foundCells = Nothing
    Set rngTarget = Nothing
    Set wb1 = Nothing
    Set wb2 = Nothing
    Set ws1 = Nothing
    Set ws2 = Nothing
    
    MsgBox "Le traitement est compl�t�", vbInformation
           
End Sub
