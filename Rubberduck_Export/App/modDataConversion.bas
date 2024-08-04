Attribute VB_Name = "modDataConversion"
Option Explicit

'Importation des clients à partir de ... \DataConversion\Clients.xlsx
Sub Copy_Data_Between_Closed_Workbooks_Clients() '2024-08-03 @ 09:40

    Dim sourceRange As Range
    
    'Définir les chemins d'accès des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALITÉ\DataConversion\Clients.xlsx"
    Dim destinationFilePath As String
    destinationFilePath = ThisWorkbook.Worksheets("Admin").Range("FolderSharedData") _
                          .value & Application.PathSeparator & "GCF_BD_Entrée.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("Feuil1")
    
    'Détermine la dernière rangée utilisée dans le fichier Source
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
    
    MsgBox "Les données ont été copiées avec succès dans le fichier destination."
    
End Sub

'Ajustements à la feuille DB_Clients (*) ---> [*]
Sub Clients_Process_Rows()

    Dim i As Long
    
    'Define the worksheet you want to work with
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BD_Clients")
    
    'Find the last used row with data in column A
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    
    'Loop through each row starting from row 2 (headers are 1 row)
    Dim client As String, client_ID As String, contactFacturation As String
    Dim posOP As Integer, posEP As Integer
    For i = 2 To lastUsedRow
        'Read data from the current row
        client = ws.Cells(i, 1).value
        client_ID = ws.Cells(i, 2).value
        contactFacturation = ws.Cells(i, 3).value
        
        'Process the data and make adjustments if necessary
        posOP = InStr(client, "(")
        posEP = InStr(client, ")")
        If posOP <> 0 Or posEP <> 0 Then
            If posEP > posOP + 5 Then
                Debug.Print i & " - " & client
            End If
        End If
        
    Next i
    
    MsgBox "Le traitement est complété sur " & i - 1 & " lignes"
    
End Sub

Sub Import_Data_From_Closed_Workbooks_TEC() '2024-08-03 @ 16:15

    Call Client_List_Import_All
    
    Dim strConnection As String
    Dim wsDest As Worksheet
    Dim i As Long, j As Long
    Dim lastUsedRow As Long
    Dim rowNum As Long
    
    'Define the path to the closed workbook
    Dim strFilePath As String
    strFilePath = "C:\VBA\GC_FISCALITÉ\DataConversion\TEC.xlsx"
    Dim strSheetName As String
    strSheetName = "TEC$"
    Dim strRange As String
    strRange = "A1:F262" 'Adjust the range as needed
    
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
    Set wsDest = ThisWorkbook.Sheets("TEC_Local")
    
    'Get the last row in the destination sheet
    lastUsedRow = wsDest.Range("A999").End(xlUp).row
    rowNum = lastUsedRow
    
    'Loop through the recordset and write data to the destination sheet
    Dim prof As String
    Dim client As String
    Dim clientCode As String
    Dim clientCodeFromDB As String
    Dim errorMesg As String
    Dim TEC_ID As Long
    Dim totHres As Double
    Do Until rst.EOF
        rowNum = rowNum + 1
        prof = rst.Fields(0).value
        client = rst.Fields(2).value
        clientCode = Left(client, 10)
            clientCode = Left(clientCode, InStr(clientCode, " -") - 1)
        client = Mid(client, InStr(client, " - ") + 3, Len(client))
        totHres = totHres + rst.Fields(4)
        
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
        
        TEC_ID = TEC_ID + 1
        wsDest.Range("A" & rowNum).value = TEC_ID
        wsDest.Range("B" & rowNum).value = Get_ID_From_Initials(prof)
        wsDest.Range("C" & rowNum).value = prof
        wsDest.Range("D" & rowNum).value = rst.Fields(1).value
        wsDest.Range("E" & rowNum).value = clientCodeFromDB
        wsDest.Range("F" & rowNum).value = client
        wsDest.Range("G" & rowNum).value = rst.Fields(3).value
        wsDest.Range("H" & rowNum).value = rst.Fields(4).value
        wsDest.Range("I" & rowNum).value = rst.Fields(5).value
        wsDest.Range("J" & rowNum).value = "VRAI"
        wsDest.Range("K" & rowNum).value = "08/03/2024"
        wsDest.Range("L" & rowNum).value = "FAUX"
        wsDest.Range("M" & rowNum).value = ""
        wsDest.Range("N" & rowNum).value = "FAUX"
        wsDest.Range("O" & rowNum).value = APP_VERSION_NO
        wsDest.Range("P" & rowNum).value = ""
        
        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les TEC ont été importés, pour un total de " & totHres & " heures"
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

'Importation des fournisseurs à partir de ... \DataConversion\Fournisseurs.xlsx
Sub Import_Data_From_Closed_Workbooks_Fournisseurs() '2024-08-03 @ 18:10

    'Définir les chemins d'accès des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALITÉ\DataConversion\Fournisseurs.xlsx"
    Dim destinationFilePath As String
    destinationFilePath = ThisWorkbook.Worksheets("Admin").Range("FolderSharedData") _
                          .value & Application.PathSeparator & "GCF_BD_Entrée.xlsx"
    
    'Declare le Workbook & le Worksheet (source)
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Dim sourceSheet As Worksheet: Set sourceSheet = sourceWorkbook.Worksheets("Fournisseurs")
    
    'Détermine la dernière rangée utilisée dans le fichier Source
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
    
    MsgBox "Les données (fournisseurs) ont été copiées avec succès dans" & vbNewLine & _
            vbNewLine & "le fichier destination."
    
End Sub

Sub Import_Data_From_Closed_Workbooks_GL_BV() '2024-08-03 @ 18:20

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
    Dim No_Entrée As Long
    No_Entrée = 1
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
        wsDest.Range("J" & rowNum).value = Format$(Now(), "mm/dd/yyyy hh:mm:ss")
        
        rst.MoveNext
        
    Loop
    
    If errorMesg <> "" Then
        MsgBox errorMesg
    Else
        MsgBox "Tous les soldes ont été importés, pour un total débit de " & totalDT & _
                vbNewLine & vbNewLine & "un total crédit de " & totalCT
    End If
    
    'Clean up
    rst.Close
    cnn.Close
    
    Set rst = Nothing
    Set cnn = Nothing
    
End Sub


