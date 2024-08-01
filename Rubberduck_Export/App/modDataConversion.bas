Attribute VB_Name = "modDataConversion"
Option Explicit

Sub Copy_Data_Between_Closed_Workbooks() '2024-07-31 @ 17:38

    Dim sourceRange As Range
    
    'D�finir les chemins d'acc�s des fichiers (source & destination)
    Dim sourceFilePath As String
    sourceFilePath = "C:\VBA\GC_FISCALIT�\DataConversion\Clients.xlsx"
    Dim destinationFilePath As String
    destinationFilePath = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    
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
    
    MsgBox "Data copied successfully from source to destination workbook."
    
End Sub

Sub Import_Clients_Into_GCF_BD_Entr�e()

    'Cette proc�dure effectue les choses suivantes:
    '   - Enl�ve les lignes actuelles du fichier GCF_BD_Entr�e.xlsx
    '   - Avec ADO, lit le fichier source de la conversion
    '               �crit toutes les lignes du fichier source dans GCF_BD_Entr�e.xlsx
    
    Call Clear_Rows_Except_Headers
    
    Dim sourcePath As String
    Dim destPath As String
    Dim connSource As Object
    Dim connDest As Object
    Dim rsSource As Object
    Dim rsDest As Object
    Dim sourceSheetName As String
    Dim destSheetName As String
    Dim sourceColumnNames As Collection
    Dim destColumnNames As Collection
    Dim i As Integer
    Dim sourceColumnName As Variant
    Dim destColumnName As Variant
    Dim colMap As Object
    
    'D�finir les chemins des workbooks source et destination
    sourcePath = "C:\VBA\GC_FISCALIT�\DataConversion\Clients.xlsx"
    destPath = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    
    'Nom de la feuille � importer (doit �tre le m�me dans les deux workbooks)
    sourceSheetName = "Feuil1"
    destSheetName = "Clients"
    
    'Cr�er des objets Connection
    Set connSource = CreateObject("ADODB.Connection")
    Set connDest = CreateObject("ADODB.Connection")
    
    'Ouvrir les connexions vers les workbooks source et destination
    connSource.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    connDest.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destPath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    
    'Efface les donn�es existantes dans le fichier destination
    
    'Lire les colonnes du workbook source
    Set sourceColumnNames = New Collection
    Set rsSource = connSource.Execute("SELECT * FROM [" & sourceSheetName & "$] WHERE 1=0") 'R�cup�rer les colonnes sans donn�es
    For i = 0 To rsSource.Fields.count - 1
        sourceColumnNames.add rsSource.Fields(i).name
    Next i
    rsSource.Close
    
    'Lire les colonnes du workbook destination
    Set destColumnNames = New Collection
    Set rsDest = connDest.Execute("SELECT * FROM [" & destSheetName & "$] WHERE 1=0")
    For i = 0 To rsDest.Fields.count - 1
        destColumnNames.add rsDest.Fields(i).name
    Next i
    rsDest.Close
    
    'Mapper les colonnes du workbook source aux colonnes du workbook destination
    Set colMap = CreateObject("Scripting.Dictionary")
    For Each sourceColumnName In sourceColumnNames
        For Each destColumnName In destColumnNames
            If LCase(sourceColumnName) = LCase(destColumnName) Then
                colMap.add sourceColumnName, destColumnName
                Exit For
            End If
        Next destColumnName
    Next sourceColumnName
    
    'Lire les donn�es du workbook source et ins�rer dans le workbook destination
    Set rsSource = connSource.Execute("SELECT * FROM [" & sourceSheetName & "$]")
    Do While Not rsSource.EOF
        Dim insertQuery As String
        insertQuery = "INSERT INTO [" & destSheetName & "$] ("
        
        'Construire les colonnes de la requ�te d'insertion
        For Each sourceColumnName In colMap.keys
            insertQuery = insertQuery & "[" & colMap(sourceColumnName) & "],"
        Next sourceColumnName
        insertQuery = Left(insertQuery, Len(insertQuery) - 1) & ") VALUES ("
        
        'Construire les valeurs de la requ�te d'insertion
        For Each sourceColumnName In colMap.keys
            insertQuery = insertQuery & "'" & IIf(rsSource.Fields(sourceColumnName).value = "", "NULL", rsSource.Fields(sourceColumnName).value & "'") & ","
        Next sourceColumnName
        insertQuery = Left(insertQuery, Len(insertQuery) - 1) & ")"
        
        'Ex�cuter la requ�te d'insertion
'        On Error Resume Next ' Ignorer les erreurs li�es aux feuilles prot�g�es ou non modifiables
        
        Debug.Print insertQuery
        connDest.Execute insertQuery
'        On Error GoTo 0
        
        'Passer � la ligne suivante
        rsSource.MoveNext
    Loop
    
    rsSource.Close

    'Fermer les connexions
    connSource.Close
    connDest.Close
    
    'Lib�rer la m�moire
    Set connSource = Nothing
    Set connDest = Nothing
    
    MsgBox "Les donn�es ont �t� import�es avec succ�s !"
    
End Sub

Sub Clear_Rows_Except_Headers()

    'Chemin du fichier � ouvrir
    Dim fullFileName As String
    fullFileName = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    
    'Ouvrir le fichier .xlsx
    Dim wb As Workbook: Set wb = Workbooks.Open(fullFileName)
    
    'D�finir la feuille � traiter
    Dim ws As Worksheet: Set ws = wb.Worksheets("Clients")
    
    'Trouver la derni�re ligne utilis�e dans la feuille
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    
    'Effacer toutes les lignes sauf la premi�re (ligne d'en-t�te)
    If lastUsedRow > 1 Then
        ws.rows("2:" & lastUsedRow).ClearContents 'Formats remain
    End If
    
    'Sauvegarder et fermer le fichier
    wb.Save
    wb.Close
    
    'Nettoyer les objets
    Set ws = Nothing
    Set wb = Nothing
    
End Sub

