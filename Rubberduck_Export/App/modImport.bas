Attribute VB_Name = "modImport"
Option Explicit

Sub ChartOfAccount_Import_All() '2024-02-17 @ 07:21

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ChartOfAccount_Import_All", "", 0)

    'Clear all cells, but the headers, in the target worksheet
    wshAdmin.Range("T10").CurrentRegion.offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx"
    sourceWorksheet = "PlanComptable$"

    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection

    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open

    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceWorksheet & "]"
        .Open
    End With

    'Copy to wshAdmin workbook
    wshAdmin.Range("T11").CopyFromRecordset recSet

    'Close resource
    recSet.Close
    connStr.Close

    Call Dynamic_Range_Redefine_Plan_Comptable

    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing

    Call Log_Record("modImport:ChartOfAccount_Import_All", "", startTime)

End Sub

Sub Client_List_Import_All() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Client_List_Import_All", "", 0)
    
'    Application.ScreenUpdating = False
    
    'Worksheet recevant les données importées
    Dim wsLocal As Worksheet: Set wsLocal = wshBD_Clients
    
    'Efface toutes les lignes, sauf la ligne d'entête
    wsLocal.Range("A1").CurrentRegion.offset(1, 0).ClearContents
    
    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx" '2024-02-14 @ 07:04
    sourceTab = "Clients$"
    
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        recSet.Open
    End With
    
    'Copier le recSet vers wsLocal
    wsLocal.Range("A2").CopyFromRecordset recSet
    
    'Redimensionner le tableau & appliquer le format
    Dim tableName As String
    tableName = "l_tbl_BD_Clients"
    
    Call ResizeTable(wsLocal, tableName)
    Call ApplyFormatting(wsLocal, tableName)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set wsLocal = Nothing
    
    Call Log_Record("modImport:Client_List_Import_All", "", startTime)

End Sub

Sub DEB_Récurrent_Import_All() '2024-07-08 @ 08:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:DEB_Récurrent_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshDEB_Récurrent.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx', in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "DEB_Récurrent$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    'Copy to wshDEB_Récurrent workbook
    wshDEB_Récurrent.Range("A2").CopyFromRecordset recSet

    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshDEB_Récurrent.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshDEB_Récurrent, rng, 1)
    
    Call DEB_Recurrent_Build_Summary '2025-01-15 @
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:DEB_Récurrent_Import_All", "", startTime)

End Sub

Sub DEB_Trans_Import_All() '2024-06-26 @ 18:51
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:DEB_Trans_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshDEB_Trans.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "DEB_Trans$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshDEB_Trans workbook after erasing actual lines
    wshDEB_Trans.Rows("2:" & wshDEB_Trans.Rows.count).ClearContents
    wshDEB_Trans.Range("A2").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshDEB_Trans.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshDEB_Trans, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:DEB_Trans_Import_All", "", startTime)

End Sub

Sub ENC_Détails_Import_All() '2025-01-16 @ 16:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ENC_Détails_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_Détails.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "ENC_Détails$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshENC_Détails workbook
    wshENC_Détails.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshENC_Détails.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshENC_Détails, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ENC_Détails_Import_All", "", startTime)

End Sub

Sub ENC_Entête_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ENC_Entête_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_Entête.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "ENC_Entête$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshENC_Entête workbook
    wshENC_Entête.Range("A2").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:36
    Dim rng As Range: Set rng = wshENC_Entête.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshENC_Entête, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ENC_Entête_Import_All", "", startTime)

End Sub

Sub CC_Régularisations_Import_All() '2025-01-05 @ 11:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:CC_Régularisations_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshCC_Régularisations.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "CC_Régularisations$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshCC_Régularisations worksheet
    If recSet.EOF = False Then
        wshCC_Régularisations.Range("A2").CopyFromRecordset recSet
    End If

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshCC_Régularisations.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshCC_Régularisations, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:CC_Régularisations_Import_All", "", startTime)

End Sub

Sub FAC_Comptes_Clients_Import_All() '2024-08-07 @ 17:41
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Comptes_Clients_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Comptes_Clients.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import FAC_Comptes_Clients from 'GCF_DB_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "FAC_Comptes_Clients$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshCAR workbook
    wshFAC_Comptes_Clients.Range("A3").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshFAC_Comptes_Clients.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshFAC_Comptes_Clients, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Comptes_Clients_Import_All", "", startTime)

End Sub

Sub FAC_Détails_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Détails_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Détails.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Détails$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshFAC_Détails workbook
    wshFAC_Détails.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshFAC_Détails.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshFAC_Détails, rng, 2)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Détails_Import_All", "", startTime)

End Sub

Sub FAC_Entête_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Entête_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Entête.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Entête$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshFAC_Entête workbook
    wshFAC_Entête.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Entête.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshFAC_Entête, rng, 2)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Entête_Import_All", "", startTime)

End Sub

Sub FAC_Sommaire_Taux_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Sommaire_Taux_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Sommaire_Taux.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import FAC_Sommaire_Taux from 'GCF_BD_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Sommaire_Taux$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshFAC_Entête workbook
    wshFAC_Sommaire_Taux.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Sommaire_Taux.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshFAC_Sommaire_Taux, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Sommaire_Taux_Import_All", "", startTime)

End Sub

Sub FAC_Projets_Détails_Import_All() '2024-07-20 @ 13:25
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_Détails_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import FAC_Projets_Détails from 'GCF_DB_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Projets_Détails$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    
    'Définir le type de curseur pour permettre l'utilisation de .RecordCount - 2024-11-08 @ 06:45 - RMV
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy all rows to wshFAC_Projets_Détails workbook
    If recSet.RecordCount > 0 Then
        ws.Range("A2").CopyFromRecordset recSet
    End If

    Dim dataRange As Range
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 1 Then
        Set dataRange = ws.Range("A2:A" & lastRow)
    End If
    
    'Enlever les lignes qui doivent être enlevées
    With ws
        .Range("A1").AutoFilter Field:=9, Criteria1:="=VRAI", Operator:=xlOr, Criteria2:=-1
        On Error Resume Next
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).Delete
        On Error GoTo 0
        .AutoFilterMode = False
    End With
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 1 Then
        Set dataRange = ws.Range("A1").CurrentRegion
        Call ApplyWorksheetFormat(wshFAC_Projets_Détails, dataRange, 1)
    End If
    
    'Libérer la mémoire
    If Not connStr Is Nothing Then
        If connStr.state = adStateOpen Then connStr.Close
        Set connStr = Nothing
    End If
    If Not recSet Is Nothing Then
        If recSet.state = adStateOpen Then recSet.Close
        Set recSet = Nothing
    End If
    Set dataRange = Nothing
    Set ws = Nothing
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Call Log_Record("modImport:FAC_Projets_Détails_Import_All", "", startTime)

End Sub

'Sub FAC_Projets_Détails_Import_All_OK() '2024-07-20 @ 13:25
'
'    Dim startTime as Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_Détails_Import_All", "", 0)
'
'    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
'
'    'Clear all cells, but the headers, in the target worksheet
'    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents
'
'    'Import FAC_Projets_Détails from 'GCF_DB_MASTER.xlsx'
'    Dim sourceWorkbook As String, sourceTab As String
'    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
'                     "GCF_BD_MASTER.xlsx"
'    sourceTab = "FAC_Projets_Détails$"
'
'    'ADODB connection
'    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
'
'    'Connection String specific to EXCEL
'    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
'                               "Data Source = " & sourceWorkbook & ";" & _
'                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
'    connStr.Open
'
'    'Recordset
'    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
'
'    'Définir le type de curseur pour permettre l'utilisation de .RecordCount - 2024-11-08 @ 06:45 - RMV
'    With recSet
'        .ActiveConnection = connStr
'        .CursorType = adOpenStatic
'        .LockType = adLockReadOnly
'        .source = "SELECT * FROM [" & sourceTab & "]"
'        .Open
'    End With
'
'    'Copy all rows to wshFAC_Projets_Détails workbook
'    If recSet.RecordCount > 0 Then
'        ws.Range("A2").CopyFromRecordset recSet
'    End If
'
'    Dim lastRow As Long
'    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
'
'    'Enlever les lignes dont la valeur estDétruite est égale à VRAI / -1 - 2024-11-15 @ 07:53
'    Dim i As Long
'    For i = lastRow To 2 Step -1
'        If UCase(ws.Cells(i, "I")) = "VRAI" Or ws.Cells(i, "I") = -1 Then
'            ws.Rows(i).Delete
'        End If
'    Next i
'
'   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
'    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
'    If lastRow > 1 Then
'        Dim rng As Range: Set rng = wshFAC_Projets_Détails.Range("A1").CurrentRegion
'        Call ApplyWorksheetFormat(wshFAC_Projets_Détails, rng, 1)
'    End If
'
'    Application.ScreenUpdating = True
'
'    'Libérer la mémoire
'    Set connStr = Nothing
'    Set recSet = Nothing
'    Set rng = Nothing
'    Set ws = Nothing
'
'    Call Log_Record("modImport:FAC_Projets_Détails_Import_All", "", startTime)
'
'End Sub
'
Sub FAC_Projets_Entête_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_Entête_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Projets_Entête$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    
    'Définir le type de curseur pour permettre l'utilisation de .RecordCount
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshFAC_Projets_Entête workbook
    If recSet.RecordCount > 0 Then
        ws.Range("A2").CopyFromRecordset recSet
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    'Delete the rows that column (isDétruite) is set to TRUE
    Dim i As Long
    If lastRow >= 2 Then
        For i = lastRow To 2 Step -1
            If UCase(ws.Cells(i, 26).Value) = "VRAI" Or _
                ws.Cells(i, 26).Value = -1 Then
                ws.Rows(i).Delete
            End If
        Next i
    End If
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 1 Then
        Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
        Call ApplyWorksheetFormat(ws, rng, 1)
    End If
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:FAC_Projets_Entête_Import_All", "", startTime)

End Sub

Sub Fournisseur_List_Import_All() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Fournisseur_List_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBD_Fournisseurs.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx" '2024-02-14 @ 07:04
    sourceTab = "Fournisseurs$"
    
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshBD_Fournisseurs workbook
    wshBD_Fournisseurs.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    Dim rng As Range: Set rng = wshBD_Fournisseurs.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshBD_Fournisseurs, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:Fournisseur_List_Import_All", "", startTime)

End Sub

Sub GL_EJ_Recurrente_Import_All() '2024-03-03 @ 11:36

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:GL_EJ_Recurrente_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim lastUsedRow As Long
    lastUsedRow = wshGL_EJ_Recurrente.Cells(wshGL_EJ_Recurrente.Rows.count, "C").End(xlUp).row
    
    'Clear all cells, but the headers and Columns A & B, in the target worksheet
    If lastUsedRow > 1 Then
        wshGL_EJ_Recurrente.Range("C2:I" & lastUsedRow).ClearContents
    End If
    
    'Import EJ_Auto from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "GL_EJ_Récurrente$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshGL_EJ_Recurrente workbook
    wshGL_EJ_Recurrente.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wshGL_EJ_Recurrente.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshGL_EJ_Recurrente, rng, 1)
    
    Call GL_EJ_Recurrente_Build_Summary '2024-03-14 @ 07:38
    
Clean_Exit:
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:GL_EJ_Recurrente_Import_All", "", startTime)

End Sub

Sub GL_Trans_Import_All() '2024-03-03 @ 10:13
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:GL_Trans_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Worksheet recevant les données importées
    Dim wsLocal As Worksheet: Set wsLocal = wshGL_Trans
    
    'Effacer toutes les lignes, sauf la ligne d'entête
    Dim saveLastRow As Long
    saveLastRow = wsLocal.Cells(wsLocal.Rows.count, 1).End(xlUp).row
    If saveLastRow > 1 Then
        wsLocal.Range("A1").CurrentRegion.offset(1, 0).ClearContents
    End If

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "GL_Trans$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    'Copy to wsLocal workbook
    wsLocal.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wsLocal.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wsLocal, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set wsLocal = Nothing
    
    Call Log_Record("modImport:GL_Trans_Import_All", "", startTime)

End Sub

Sub TEC_Import_All()                             '2024-02-14 @ 06:19
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:TEC_Import_All", "", 0)
    
'    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    'Clear all cells, but the headers, in the destination worksheet
    ws.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"        '2024-02-14 @ 06:22
    sourceTab = "TEC_Local$"
    
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshTEC_Local workbook
    ws.Range("A3").CopyFromRecordset recSet

    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:TEC_Import_All", "", startTime)

End Sub

Sub ResizeTable(targetSheet As Worksheet, tableName As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ResizeTable", "", 0)

    Dim tbl As ListObject
    Dim lastRow As Long
    Dim lastCol As Long
    
    'Trouver le tableau
    Set tbl = targetSheet.ListObjects(tableName)
    
    'Déterminer la dernière ligne et colonne des nouvelles données
    With targetSheet
        lastRow = .Cells(.Rows.count, tbl.Range.Column).End(xlUp).row
        lastCol = .Cells(tbl.Range.row, .Columns.count).End(xlToLeft).Column
    End With
    
    'Redimensionner la plage du tableau
    tbl.Resize Range(tbl.Range.Cells(1, 1), targetSheet.Cells(lastRow, lastCol))
    
    Call Log_Record("modImport:ResizeTable", "", startTime)

End Sub

Sub ApplyFormatting(targetSheet As Worksheet, tableName As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ApplyFormatting", "", 0)
    
    Dim tbl As ListObject
    
    ' Identifier le tableau
    Set tbl = targetSheet.ListObjects(tableName)
    
    ' Appliquer un style existant au tableau
    tbl.tableStyle = "TableStyleMedium2" ' Modifier selon le style souhaité
    
    Call Log_Record("modImport:ApplyFormatting", "", startTime)
    
End Sub
