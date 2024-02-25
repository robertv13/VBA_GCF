Attribute VB_Name = "modImport"
Option Explicit

Sub Client_List_Import_All() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshClientDB.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim fullFileName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & _
                   Application.PathSeparator & "GCF_BD_Entrée.xlsx" '2024-02-14 @ 07:04
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = fullFileName
    sourceTab = "Clients"
    
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    
    recSet.ActiveConnection = connStr
    recSet.source = "SELECT * FROM [" & sourceTab & "$]"
    recSet.Open
    
    'Copy to wshClientDB workbook
    wshClientDB.Range("A2").CopyFromRecordset recSet
    wshClientDB.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
'    MsgBox _
'        Prompt:="J'ai importé un total de " & _
'            Format(wshClientDB.Range("A1").CurrentRegion.Rows.count - 1, _
'            "## ##0") & " clients", _
'        Title:="Vérification du nombre de clients", _
'        Buttons:=vbInformation

    'Free up memory - 2024-02-23
    Set connStr = Nothing
    Set recSet = Nothing

    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        Now() & "- Client_List_Import_All() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")
        
End Sub

Sub TEC_Import_All() '2024-02-14 @ 06:19
    
    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBaseHours.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim fileName As String, sourceWorkbook As String, sourceTab As String
    fileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                "GCF_BD_Sortie.xlsx" '2024-02-14 @ 06:22
    sourceWorkbook = fileName
    sourceTab = "TEC"
    
    'Set up source and destination ranges
    Dim sourceRange As Range, destinationRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange
    Set destinationRange = wshBaseHours.Range("A2")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshBaseHours.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Arrange formats on all rows
    Dim lastRow As Long
    lastRow = wshBaseHours.Range("A99999").End(xlUp).row
    
    With wshBaseHours
        .Range("A3" & ":P" & lastRow).HorizontalAlignment = xlCenter
        With .Range("F3:F" & lastRow & ",G3:G" & lastRow & ",I3:I" & lastRow & ",O3:O" & lastRow)
            .HorizontalAlignment = xlLeft
        End With
        .Range("H3:H" & lastRow).NumberFormat = "#0.00"
        .Range("K3:K" & lastRow).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End With
    
    Application.ScreenUpdating = True
    
    'Free up memory - 2024-02-23
    Set sourceRange = Nothing
    Set destinationRange = Nothing

    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        Now() & " - TEC_Import_All() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")
    
End Sub

