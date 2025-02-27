Attribute VB_Name = "modImport"
Option Explicit

Sub ChartOfAccount_Import_All() '2024-02-17 @ 07:21

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ChartOfAccount_Import_All", "", 0)

    'Clear all cells, but the headers, in the target worksheet
    wshAdmin.Range("T10").CurrentRegion.offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entr�e.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx"
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

    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing

    Call Log_Record("modImport:ChartOfAccount_Import_All", "", startTime)

End Sub

Sub Client_List_Import_All() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Client_List_Import_All", "", 0)
    
'    Application.ScreenUpdating = False
    
    'Worksheet recevant les donn�es import�es
    Dim wsLocal As Worksheet: Set wsLocal = wshBD_Clients
    
    'Efface toutes les lignes, sauf la ligne d'ent�te
    wsLocal.Range("A1").CurrentRegion.offset(1, 0).ClearContents
    
    'Import Clients List from 'GCF_BD_Entr�e.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx" '2024-02-14 @ 07:04
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
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set wsLocal = Nothing
    
    Call Log_Record("modImport:Client_List_Import_All", "", startTime)

End Sub

Sub DEB_R�current_Import_All() '2024-07-08 @ 08:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:DEB_R�current_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshDEB_R�current.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx', in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "DEB_R�current$"
                     
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
    'Copy to wshDEB_R�current workbook
    wshDEB_R�current.Range("A2").CopyFromRecordset recSet

    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshDEB_R�current.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshDEB_R�current, rng, 1)
    
    Call DEB_Recurrent_Build_Summary '2025-01-15 @
    
    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:DEB_R�current_Import_All", "", startTime)

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
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:DEB_Trans_Import_All", "", startTime)

End Sub

Sub ENC_D�tails_Import_All() '2025-01-16 @ 16:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ENC_D�tails_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_D�tails.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "ENC_D�tails$"
                     
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
    
    'Copy to wshENC_D�tails workbook
    wshENC_D�tails.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshENC_D�tails.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshENC_D�tails, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ENC_D�tails_Import_All", "", startTime)

End Sub

Sub ENC_Ent�te_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ENC_Ent�te_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_Ent�te.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "ENC_Ent�te$"
                     
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
    
    'Copy to wshENC_Ent�te workbook
    wshENC_Ent�te.Range("A2").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:36
    Dim rng As Range: Set rng = wshENC_Ent�te.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshENC_Ent�te, rng, 1)

    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ENC_Ent�te_Import_All", "", startTime)

End Sub

Sub CC_R�gularisations_Import_All() '2025-01-05 @ 11:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:CC_R�gularisations_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshCC_R�gularisations.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "CC_R�gularisations$"
                     
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
    
    'Copy to wshCC_R�gularisations worksheet
    If recSet.EOF = False Then
        wshCC_R�gularisations.Range("A2").CopyFromRecordset recSet
    End If

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshCC_R�gularisations.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshCC_R�gularisations, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:CC_R�gularisations_Import_All", "", startTime)

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
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Comptes_Clients_Import_All", "", startTime)

End Sub

Sub FAC_D�tails_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_D�tails_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_D�tails.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_D�tails$"
                     
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
    
    'Copy to wshFAC_D�tails workbook
    wshFAC_D�tails.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshFAC_D�tails.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshFAC_D�tails, rng, 2)

    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_D�tails_Import_All", "", startTime)

End Sub

Sub FAC_Ent�te_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Ent�te_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Ent�te.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Ent�te$"
                     
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
    
    'Copy to wshFAC_Ent�te workbook
    wshFAC_Ent�te.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Ent�te.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshFAC_Ent�te, rng, 2)
    
    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Ent�te_Import_All", "", startTime)

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
    
    'Copy to wshFAC_Ent�te workbook
    wshFAC_Sommaire_Taux.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Sommaire_Taux.Range("A1").CurrentRegion
    Call ApplyWorksheetFormat(wshFAC_Sommaire_Taux, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Sommaire_Taux_Import_All", "", startTime)

End Sub

Sub FAC_Projets_D�tails_Import_All() '2024-07-20 @ 13:25
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_D�tails_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim ws As Worksheet: Set ws = wshFAC_Projets_D�tails
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import FAC_Projets_D�tails from 'GCF_DB_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Projets_D�tails$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    
    'D�finir le type de curseur pour permettre l'utilisation de .RecordCount - 2024-11-08 @ 06:45 - RMV
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy all rows to wshFAC_Projets_D�tails workbook
    If recSet.RecordCount > 0 Then
        ws.Range("A2").CopyFromRecordset recSet
    End If

    Dim dataRange As Range
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 1 Then
        Set dataRange = ws.Range("A2:A" & lastRow)
    End If
    
    'Enlever les lignes qui doivent �tre enlev�es
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
        Call ApplyWorksheetFormat(wshFAC_Projets_D�tails, dataRange, 1)
    End If
    
    'Lib�rer la m�moire
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
    
    Call Log_Record("modImport:FAC_Projets_D�tails_Import_All", "", startTime)

End Sub

'Sub FAC_Projets_D�tails_Import_All_OK() '2024-07-20 @ 13:25
'
'    Dim startTime as Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_D�tails_Import_All", "", 0)
'
'    Dim ws As Worksheet: Set ws = wshFAC_Projets_D�tails
'
'    'Clear all cells, but the headers, in the target worksheet
'    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents
'
'    'Import FAC_Projets_D�tails from 'GCF_DB_MASTER.xlsx'
'    Dim sourceWorkbook As String, sourceTab As String
'    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
'                     "GCF_BD_MASTER.xlsx"
'    sourceTab = "FAC_Projets_D�tails$"
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
'    'D�finir le type de curseur pour permettre l'utilisation de .RecordCount - 2024-11-08 @ 06:45 - RMV
'    With recSet
'        .ActiveConnection = connStr
'        .CursorType = adOpenStatic
'        .LockType = adLockReadOnly
'        .source = "SELECT * FROM [" & sourceTab & "]"
'        .Open
'    End With
'
'    'Copy all rows to wshFAC_Projets_D�tails workbook
'    If recSet.RecordCount > 0 Then
'        ws.Range("A2").CopyFromRecordset recSet
'    End If
'
'    Dim lastRow As Long
'    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
'
'    'Enlever les lignes dont la valeur estD�truite est �gale � VRAI / -1 - 2024-11-15 @ 07:53
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
'        Dim rng As Range: Set rng = wshFAC_Projets_D�tails.Range("A1").CurrentRegion
'        Call ApplyWorksheetFormat(wshFAC_Projets_D�tails, rng, 1)
'    End If
'
'    Application.ScreenUpdating = True
'
'    'Lib�rer la m�moire
'    Set connStr = Nothing
'    Set recSet = Nothing
'    Set rng = Nothing
'    Set ws = Nothing
'
'    Call Log_Record("modImport:FAC_Projets_D�tails_Import_All", "", startTime)
'
'End Sub
'
Sub FAC_Projets_Ent�te_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_Ent�te_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Ent�te
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Projets_Ent�te$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    
    'D�finir le type de curseur pour permettre l'utilisation de .RecordCount
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wshFAC_Projets_Ent�te workbook
    If recSet.RecordCount > 0 Then
        ws.Range("A2").CopyFromRecordset recSet
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    'Delete the rows that column (isD�truite) is set to TRUE
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
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:FAC_Projets_Ent�te_Import_All", "", startTime)

End Sub

Sub Fournisseur_List_Import_All() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Fournisseur_List_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBD_Fournisseurs.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entr�e.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx" '2024-02-14 @ 07:04
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
    
    'Lib�rer la m�moire
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
    sourceTab = "GL_EJ_R�currente$"
                     
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
    
    'Lib�rer la m�moire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:GL_EJ_Recurrente_Import_All", "", startTime)

End Sub

Sub GL_Trans_Import_All() '2024-03-03 @ 10:13
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:GL_Trans_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Worksheet recevant les donn�es import�es
    Dim wsLocal As Worksheet: Set wsLocal = wshGL_Trans
    
    'Effacer toutes les lignes, sauf la ligne d'ent�te
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
    
    'Lib�rer la m�moire
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

    'Lib�rer la m�moire
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
    
    'D�terminer la derni�re ligne et colonne des nouvelles donn�es
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
    tbl.tableStyle = "TableStyleMedium2" ' Modifier selon le style souhait�
    
    Call Log_Record("modImport:ApplyFormatting", "", startTime)
    
End Sub
