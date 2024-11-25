Attribute VB_Name = "modImport"
Option Explicit

'CommentOut - 2024-11-20 @ 07:11
'Sub Admin_Import_Worksheet() '2024-07-02 @ 10:14
'
'    'Save the shared data folder name
'    Dim saveDataPath As String
'    saveDataPath = wshAdmin.Range("F5").value & DATA_PATH
'
'    'Define the target workbook and sheet names
'    Dim targetWorkbook As Workbook: Set targetWorkbook = ThisWorkbook
'    Dim targetSheetName As String
'    targetSheetName = "Admin"
'    Dim sourceSheetName As String
'    sourceSheetName = "Admin_Master"
'
'    'Open the source workbook
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
'    Dim sourceWorkbook As Workbook: Set sourceWorkbook = _
'        Workbooks.Open(saveDataPath & Application.PathSeparator & "GCF_BD_MASTER.xlsx")
'
'    Debug.Print "#066 - Source     : " & sourceWorkbook.Name & " with " & sourceSheetName
'    Debug.Print "#067 - Destination: " & targetWorkbook.Name & " with " & targetSheetName
'
'    'Copy the source worksheet
'    sourceWorkbook.Sheets(sourceSheetName).Copy Before:=targetWorkbook.Sheets(2)
'    Debug.Print "#068 - The new sheet is created..."
'    Dim tempSheet As Worksheet: Set tempSheet = targetWorkbook.Sheets(2)
'    tempSheet.Name = "TempSheetName"
'    Debug.Print "#069 - The new sheet is now called 'TempSheetName'"
'
'    'Delete the old (target) worksheet
'    Debug.Print "#070 - About to delete '" & targetSheetName & "'"
'    targetWorkbook.Sheets(targetSheetName).Delete
'
'    'Rename the copied worksheet to the target worksheet name
'    tempSheet.Name = targetSheetName
'
''    'Change the code name of the worksheet
''    Dim vbaProject As Object: Set vbaProject = targetWorkbook.VBProject
''    Dim vbaComponent As Object: Set vbaComponent = vbaProject.VBComponents("Feuil2")
''    vbaComponent.Properties("_CodeName").value = "wshADMIN"
'
'    'Close the source workbook
'    sourceWorkbook.Close SaveChanges:=False
'    Application.DisplayAlerts = True
'    Application.ScreenUpdating = True
'
'    'Cleaning - 2024-07-02 @ 14:27
'    Set sourceWorkbook = Nothing
'    Set targetWorkbook = Nothing
'    Set tempSheet = Nothing
'
'End Sub
'
Sub ChartOfAccount_Import_All() '2024-02-17 @ 07:21

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ChartOfAccount_Import_All", 0)

    'Clear all cells, but the headers, in the target worksheet
    wshAdmin.Range("T10").CurrentRegion.offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceWorksheet & "]"
    recSet.Open

    'Copy to wshAdmin workbook
    wshAdmin.Range("T11").CopyFromRecordset recSet

    'Close resource
    recSet.Close
    connStr.Close

    Call Dynamic_Range_Redefine_Plan_Comptable

    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing

    Call Log_Record("modImport:ChartOfAccount_Import_All", startTime)

End Sub

Sub Client_List_Import_All() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Client_List_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Worksheet recevant les données importées
    Dim wsLocal As Worksheet: Set wsLocal = wshBD_Clients
    
    'Efface toutes les lignes, sauf la ligne d'entête
    wsLocal.Range("A1").CurrentRegion.offset(1, 0).ClearContents
    
    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copier le recSet vers wsLocal
    wsLocal.Range("A2").CopyFromRecordset recSet
    
    'Redimensionner le tableau local
    Dim tblLocal As ListObject
    Dim dataRange As Range
    Dim lastRow As Long, lastCol As Long
    If wsLocal.ListObjects.count > 0 Then
        Set tblLocal = wsLocal.ListObjects(1) 'Supposons qu'il y a un seul tableau par feuille
        'Rechercher la dernière ligne avec des données
        Dim i As Long
        For i = tblLocal.ListRows.count To 1 Step -1
            If tblLocal.DataBodyRange.Cells(i, 2).value <> "" Then
                lastRow = i
                Exit For
            End If
        Next i
'        lastRow = tblLocal.ListRows.count
        lastCol = wsLocal.Cells(1, wsLocal.Columns.count).End(xlToLeft).Column
        Set dataRange = wsLocal.Range("A1", wsLocal.Cells(lastRow + 1, lastCol))
        'Redimensionner le tableau pour s’adapter à la nouvelle plage
        tblLocal.Resize dataRange
    End If
    
    'Setup the format of the worksheet - 2024-07-20 @ 18:31
    Dim rng As Range: Set rng = wsLocal.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wsLocal, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set dataRange = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set tblLocal = Nothing
    Set wsLocal = Nothing
    
    Call Log_Record("modImport:Client_List_Import_All", startTime)

End Sub

Sub DEB_Recurrent_Import_All() '2024-07-08 @ 08:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:DEB_Recurrent_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshDEB_Recurrent.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx', in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "DEB_Recurrent$"
                     
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
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshDEB_Recurrent workbook
    wshDEB_Recurrent.Range("A2").CopyFromRecordset recSet

    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshDEB_Recurrent.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshDEB_Recurrent, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:DEB_Recurrent_Import_All", startTime)

End Sub

Sub DEB_Trans_Import_All() '2024-06-26 @ 18:51
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:DEB_Trans_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshDEB_Trans.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshDEB_Trans workbook
    wshDEB_Trans.Range("A2").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshDEB_Trans.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshDEB_Trans, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:DEB_Trans_Import_All", startTime)

End Sub

Sub ENC_Détails_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ENC_Détails_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_Détails.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshENC_Détails workbook
    wshENC_Détails.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshENC_Détails.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshENC_Détails, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ENC_Détails_Import_All", startTime)

End Sub

Sub ENC_Entête_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ENC_Entête_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_Entête.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshENC_Entête workbook
    wshENC_Entête.Range("A2").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:36
    Dim rng As Range: Set rng = wshENC_Entête.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshENC_Entête, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ENC_Entête_Import_All", startTime)

End Sub

Sub FAC_Comptes_Clients_Import_All() '2024-08-07 @ 17:41
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Comptes_Clients_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Comptes_Clients.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import FAC_Comptes_Clients from 'GCF_DB_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshCAR workbook
    wshFAC_Comptes_Clients.Range("A3").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshFAC_Comptes_Clients.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Comptes_Clients, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Comptes_Clients_Import_All", startTime)

End Sub

Sub FAC_Détails_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Détails_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Détails.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshFAC_Détails workbook
    wshFAC_Détails.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshFAC_Détails.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Détails, rng, 2)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Détails_Import_All", startTime)

End Sub

Sub FAC_Entête_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Entête_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Entête.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshFAC_Entête workbook
    wshFAC_Entête.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Entête.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Entête, rng, 2)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Entête_Import_All", startTime)

End Sub

Sub FAC_Sommaire_Taux_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Sommaire_Taux_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Sommaire_Taux.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import FAC_Sommaire_Taux from 'GCF_BD_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshFAC_Entête workbook
    wshFAC_Sommaire_Taux.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Sommaire_Taux.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Entête, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Sommaire_Taux_Import_All", startTime)

End Sub

Sub FAC_Projets_Détails_Import_All() '2024-07-20 @ 13:25
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_Détails_Import_All", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import FAC_Projets_Détails from 'GCF_DB_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.CursorType = adOpenKeyset
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy all rows to wshFAC_Projets_Détails workbook
    If recSet.RecordCount > 0 Then
        ws.Range("A2").CopyFromRecordset recSet
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    'Enlever les lignes dont la valeur estDétruite est égale à VRAI / -1 - 2024-11-15 @ 07:53
    Dim i As Long
    For i = lastRow To 2 Step -1
        If UCase(ws.Cells(i, "I")) = "VRAI" Or ws.Cells(i, "I") = -1 Then
            ws.Rows(i).Delete
        End If
    Next i
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 1 Then
        Dim rng As Range: Set rng = wshFAC_Projets_Détails.Range("A1").CurrentRegion
        Call Apply_Worksheet_Format(wshFAC_Projets_Détails, rng, 1)
    End If
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:FAC_Projets_Détails_Import_All", startTime)

End Sub

Sub FAC_Projets_Entête_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_Entête_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
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
            If UCase(ws.Cells(i, 26).value) = "VRAI" Or _
                ws.Cells(i, 26).value = -1 Then
                ws.Rows(i).Delete
            End If
        Next i
    End If
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 1 Then
        Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
        Call Apply_Worksheet_Format(ws, rng, 1)
    End If
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:FAC_Projets_Entête_Import_All", startTime)

End Sub

Sub Fournisseur_List_Import_All() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Fournisseur_List_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBD_Fournisseurs.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshBD_Fournisseurs workbook
    wshBD_Fournisseurs.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    Dim rng As Range: Set rng = wshBD_Fournisseurs.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshBD_Fournisseurs, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:Fournisseur_List_Import_All", startTime)

End Sub

Sub GL_EJ_Recurrente_Import_All() '2024-03-03 @ 11:36

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:GL_EJ_Recurrente_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    Dim lastUsedRow As Long
    lastUsedRow = wshGL_EJ_Recurrente.Cells(wshGL_EJ_Recurrente.Rows.count, "C").End(xlUp).row
    
    'Clear all cells, but the headers and Columns A & B, in the target worksheet
    If lastUsedRow > 1 Then
        wshGL_EJ_Recurrente.Range("C2:I" & lastUsedRow).ClearContents
    End If
    
    'Import EJ_Auto from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "GL_EJ_Recurrente$"
                     
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
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshGL_EJ_Recurrente workbook
    wshGL_EJ_Recurrente.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wshGL_EJ_Recurrente.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshGL_EJ_Recurrente, rng, 1)
    
    Call GL_EJ_Recurrente_Build_Summary '2024-03-14 @ 07:38
    
Clean_Exit:
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:GL_EJ_Recurrente_Import_All", startTime)

End Sub

Sub GL_Trans_Import_All() '2024-03-03 @ 10:13
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:GL_Trans_Import_All", 0)
    
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
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wsLocal workbook
    wsLocal.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wsLocal.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wsLocal, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set wsLocal = Nothing
    
    Call Log_Record("modImport:GL_Trans_Import_All", startTime)

End Sub

Sub TEC_Import_All()                             '2024-02-14 @ 06:19
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:TEC_Import_All", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    'Clear all cells, but the headers, in the destination worksheet
    ws.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    recSet.ActiveConnection = connStr
    recSet.CursorType = adOpenStatic
    recSet.LockType = adLockReadOnly
    recSet.source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshTEC_Local workbook
    ws.Range("A3").CopyFromRecordset recSet

    'Redimensionner le tableau
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("l_tbl_TEC_Local")
    Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
    
    'Setup the format of the worksheet using a Sub
    Call Apply_Worksheet_Format(ws, rng, 2)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set tbl = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:TEC_Import_All", startTime)

End Sub



