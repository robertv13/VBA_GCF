Attribute VB_Name = "modImport"
Option Explicit

Sub Admin_Import_Worksheet() '2024-07-02 @ 10:14
    
    Application.StatusBar = "J'importe la feuille 'Admin'"
    
    'Save the shared data folder name
    Dim saveDataPath As String
    saveDataPath = wshAdmin.Range("F5").value & DATA_PATH
    
    'Define the target workbook and sheet names
    Dim targetWorkbook As Workbook: Set targetWorkbook = ThisWorkbook
    Dim targetSheetName As String
    targetSheetName = "Admin"
    Dim sourceSheetName As String
    sourceSheetName = "Admin_Master"
    
    'Open the source workbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = _
        Workbooks.Open(saveDataPath & Application.PathSeparator & "GCF_BD_MASTER.xlsx")
    
    Debug.Print "Source     : " & sourceWorkbook.name & " with " & sourceSheetName
    Debug.Print "Destination: " & targetWorkbook.name & " with " & targetSheetName
    
    'Copy the source worksheet
    sourceWorkbook.Sheets(sourceSheetName).Copy Before:=targetWorkbook.Sheets(2)
    Debug.Print "The new sheet is created..."
    Dim tempSheet As Worksheet: Set tempSheet = targetWorkbook.Sheets(2)
    tempSheet.name = "TempSheetName"
    Debug.Print "The new sheet is now called 'TempSheetName'"

    'Delete the old (target) worksheet
    Debug.Print "About to delete '" & targetSheetName & "'"
    targetWorkbook.Sheets(targetSheetName).Delete

    'Rename the copied worksheet to the target worksheet name
    tempSheet.name = targetSheetName

'    'Change the code name of the worksheet
'    Dim vbaProject As Object: Set vbaProject = targetWorkbook.VBProject
'    Dim vbaComponent As Object: Set vbaComponent = vbaProject.VBComponents("Feuil2")
'    vbaComponent.Properties("_CodeName").value = "wshADMIN"
    
    'Close the source workbook
    sourceWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    'Cleaning - 2024-07-02 @ 14:27
    Set sourceWorkbook = Nothing
    Set targetWorkbook = Nothing
    Set tempSheet = Nothing
    
End Sub

Sub ChartOfAccount_Import_All() '2024-02-17 @ 07:21

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ChartOfAccount_Import_All", 0)
    
    Application.StatusBar = "J'importe le plan comptable"
    
    'Clear all cells, but the headers, in the target worksheet
    wshAdmin.Range("T10").CurrentRegion.Offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx"
    sourceWorksheet = "PlanComptable"
    
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
    recSet.source = "SELECT * FROM [" & sourceWorksheet & "$]"
    recSet.Open
    
    'Copy to wshAdmin workbook
    wshAdmin.Range("T11").CopyFromRecordset recSet
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Call Dynamic_Range_Redefine_Plan_Comptable
        
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    
    Call Log_Record("modImport:ChartOfAccount_Import_All()", startTime)

End Sub

Sub Client_List_Import_All() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Client_List_Import_All", 0)
    
    Application.StatusBar = "J'importe la liste des clients"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBD_Clients.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx" '2024-02-14 @ 07:04
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
    
    'Copy to wshBD_Clients workbook
    wshBD_Clients.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet - 2024-07-20 @ 18:31
    Dim rng As Range: Set rng = wshBD_Clients.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshBD_Clients, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""

    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    
    Call Log_Record("modImport:Client_List_Import_All", startTime)

End Sub

Sub DEB_Recurrent_Import_All() '2024-07-08 @ 08:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:DEB_Recurrent_Import_All", 0)
    
    Application.StatusBar = "J'importe les transactions récurrentes de déboursés"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshDEB_Recurrent.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx', in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "DEB_Recurrent"
                     
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
    
    'Copy to wshDEB_Recurrent workbook
    wshDEB_Recurrent.Range("A2").CopyFromRecordset recSet

    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshDEB_Recurrent.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshDEB_Recurrent, rng, 1)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:DEB_Recurrent_Import_All", startTime)

End Sub

Sub DEB_Trans_Import_All() '2024-06-26 @ 18:51
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:DEB_Trans_Import_All", 0)
    
    Application.StatusBar = "J'importe les transactions de déboursés"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshDEB_Trans.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "DEB_Trans"
                     
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
    
    'Copy to wshDEB_Trans workbook
    wshDEB_Trans.Range("A2").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshDEB_Trans.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshDEB_Trans, rng, 1)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:DEB_Trans_Import_All", startTime)

End Sub

Sub ENC_Détails_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ENC_Détails_Import_All", 0)
    
    Application.StatusBar = "J'importe le détail des Encaissements"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_Détails.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "ENC_Détails"
                     
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
    
    'Copy to wshENC_Détails workbook
    wshENC_Détails.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshENC_Détails.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshENC_Détails, rng, 1)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ENC_Détails_Import_All", startTime)

End Sub

Sub ENC_Entête_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ENC_Entête_Import_All", 0)
    
    Application.StatusBar = "J'importe le détail des Encaissements"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_Entête.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "ENC_Entête"
                     
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
    
    'Copy to wshENC_Entête workbook
    wshENC_Entête.Range("A2").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:36
    Dim rng As Range: Set rng = wshENC_Entête.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshENC_Entête, rng, 1)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ENC_Entête_Import_All", startTime)

End Sub

Sub FAC_Comptes_Clients_Import_All() '2024-08-07 @ 17:41
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Comptes_Clients_Import_All", 0)
    
    Application.StatusBar = "J'importe les transactions de CAR"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Comptes_Clients.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import FAC_Comptes_Clients from 'GCF_DB_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "FAC_Comptes_Clients"
                     
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
    
    'Copy to wshCAR workbook
    wshFAC_Comptes_Clients.Range("A3").CopyFromRecordset recSet
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshFAC_Comptes_Clients.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Comptes_Clients, rng, 1)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Comptes_Clients_Import_All", startTime)

End Sub

Sub FAC_Détails_Import_All() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Détails_Import_All", 0)
    
    Application.StatusBar = "J'importe le détail des Factures"
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Détails.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Détails"
                     
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
    
    'Copy to wshFAC_Détails workbook
    wshFAC_Détails.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshFAC_Détails.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Détails, rng, 2)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Détails_Import_All", startTime)

End Sub

Sub FAC_Entête_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Entête_Import_All", 0)
    
    Application.StatusBar = "J'importe les entêtes de Facture"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Entête.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Entête"
                     
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
    
    'Copy to wshFAC_Entête workbook
    wshFAC_Entête.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Entête.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Entête, rng, 2)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Entête_Import_All", startTime)

End Sub

Sub FAC_Sommaire_Taux_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Sommaire_Taux_Import_All", 0)
    
    Application.StatusBar = "J'importe les sommaires de taux"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Sommaire_Taux.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import FAC_Sommaire_Taux from 'GCF_BD_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Sommaire_Taux"
                     
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
    
    'Copy to wshFAC_Entête workbook
    wshFAC_Sommaire_Taux.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Sommaire_Taux.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Entête, rng, 1)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Sommaire_Taux_Import_All", startTime)

End Sub

Sub FAC_Projets_Détails_Import_All() '2024-07-20 @ 13:25
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_Détails_Import_All", 0)
    
    Application.StatusBar = "J'importe le détail des Projets de Factures"
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Projets_Détails"
                     
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
    
    'Copy to wshFAC_Projets_Détails workbook all rows
    ws.Range("A2").CopyFromRecordset recSet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    
    'Delete the rows that column (isDétruite) is set to TRUE in FAC_Projets_Entête
    Dim i As Long
    For i = lastRow To 2 Step -1
        If UCase(Trim(ws.Cells(i, 9).value)) = "VRAI" Or _
            Trim(ws.Cells(i, 9).value) = "" Or _
            ws.Cells(i, 9).value = -1 Then
            ws.rows(i).Delete
        End If
    Next i
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    lastRow = wshFAC_Projets_Détails.Range("A99999").End(xlUp).Row
    If lastRow > 1 Then
        Dim rng As Range: Set rng = wshFAC_Projets_Détails.Range("A1").CurrentRegion
        Call Apply_Worksheet_Format(wshFAC_Projets_Détails, rng, 1)
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-20 @ 13:30
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:FAC_Projets_Détails_Import_All", startTime)

End Sub

Sub FAC_Projets_Entête_Import_All() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:FAC_Projets_Entête_Import_All", 0)
    
    Application.StatusBar = "J'importe les entêtes de projets de facture"
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    sourceTab = "FAC_Projets_Entête"
                     
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
    
    'Copy to wshFAC_Projets_Entête workbook
    ws.Range("A2").CopyFromRecordset recSet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    
    'Delete the rows that column (isDétruite) is set to TRUE
    Dim i As Long
    For i = lastRow To 2 Step -1
        If UCase(ws.Cells(i, 26).value) = "VRAI" Or _
            ws.Cells(i, 26).value = -1 Then
            ws.rows(i).Delete
        End If
    Next i
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow > 1 Then
        Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
        Call Apply_Worksheet_Format(ws, rng, 1)
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Application.StatusBar = ""
    
    Call Log_Record("modImport:FAC_Projets_Entête_Import_All", startTime)

End Sub

Sub Fournisseur_List_Import_All() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Fournisseur_List_Import_All", 0)
    
    Application.StatusBar = "J'importe la liste des fournisseurs"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBD_Fournisseurs.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx" '2024-02-14 @ 07:04
    sourceTab = "Fournisseurs"
    
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
    
    'Copy to wshBD_Fournisseurs workbook
    wshBD_Fournisseurs.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    Dim rng As Range: Set rng = wshBD_Fournisseurs.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshBD_Fournisseurs, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
'    MsgBox _
'        Prompt:="J'ai importé un total de " & _
'            Format(wshBD_Fournisseurs.Range("A1").CurrentRegion.rows.count - 1, _
'            "##,##0") & " fournisseurs", _
'        Title:="Vérification du nombre de fournisseurs", _
'        Buttons:=vbInformation

    Application.StatusBar = ""

    'Cleaning memory - 2024-07-03 @ 15:45
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:Fournisseur_List_Import_All", startTime)

End Sub

Sub GL_EJ_Auto_Import_All() '2024-03-03 @ 11:36

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:GL_EJ_Auto_Import_All", 0)
    
    Application.StatusBar = "J'importe les écritures de journal récurrentes"
    
    Application.ScreenUpdating = False
    
    Dim lastUsedRow As Long
    lastUsedRow = wshGL_EJ_Recurrente.Range("C999").End(xlUp).Row
    
    'Clear all cells, but the headers and Columns A & B, in the target worksheet
    If lastUsedRow > 1 Then
        wshGL_EJ_Recurrente.Range("C2:I" & lastUsedRow).ClearContents
    End If
    
    'Import EJ_Auto from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "GL_EJ_Auto"
                     
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
    
    'Copy to wshGL_EJ_Recurrente workbook
    wshGL_EJ_Recurrente.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wshGL_EJ_Recurrente.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshGL_EJ_Recurrente, rng, 1)
    
    Call GL_EJ_Auto_Build_Summary '2024-03-14 @ 07:38
    
Clean_Exit:
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:GL_EJ_Auto_Import_All", startTime)

End Sub

Sub GL_Trans_Import_All() '2024-03-03 @ 10:13
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:GL_Trans_Import_All", 0)
    
    Application.StatusBar = "J'importe les transactions du Grand-Livre"
    
    Application.ScreenUpdating = False
    
    Dim saveLastRow As Long
    saveLastRow = wshGL_Trans.Range("A99999").End(xlUp).Row
    
    'Clear all cells, but the headers, in the target worksheet
    If saveLastRow > 1 Then
        wshGL_Trans.Range("A1").CurrentRegion.Offset(1, 0).ClearContents
    End If

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = "GL_Trans"
                     
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
    
    'Copy to wshGL_Trans workbook
    wshGL_Trans.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wshGL_Trans.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshGL_Trans, rng, 1)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:GL_Trans_Import_All", startTime)

End Sub

Sub TEC_Import_All() '2024-02-14 @ 06:19
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:TEC_Import_All", 0)
    
    Application.StatusBar = "J'importe tous les TEC"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshTEC_Local.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-14 @ 06:22
    sourceTab = "TEC_Local"
    
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
    
    'Copy to wshTEC_Local workbook
    wshTEC_Local.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wshTEC_Local.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshTEC_Local, rng, 2)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:TEC_Import_All", startTime)

End Sub

Sub Dynamic_Range_Redefine_Plan_Comptable() '2024-07-04 @ 10:39
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Dynamic_Range_Redefine_Plan_Comptable", 0)

    'Redefine - dnrPlanComptable_Description_Only
    'Delete existing dynamic named range (assuming it could exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptable_Description_Only").Delete
    On Error GoTo 0
    
    'Define a new dynamic named range for 'dnrPlanComptable_Description_Only'
    Dim newRangeFormula As String
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,1)"
    
    'Create the new dynamic named range
    ThisWorkbook.Names.Add name:="dnrPlanComptable_Description_Only", RefersTo:=newRangeFormula
    
    'Redefine - dnrPlanComptable_All
    'Delete existing dynamic named range (assuming it could exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptable_All").Delete
    On Error GoTo 0
    
    'Define a new dynamic named range for 'dnrPlanComptable_All'
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,4)"
    
    'Create the new dynamic named range
    ThisWorkbook.Names.Add name:="dnrPlanComptable_All", RefersTo:=newRangeFormula
    
    Call Log_Record("modImport:Dynamic_Range_Redefine_Plan_Comptable", startTime)

End Sub


