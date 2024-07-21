Attribute VB_Name = "modImport"
Option Explicit

Sub Admin_Import_Worksheet() '2024-07-02 @ 10:14
    
    Application.StatusBar = "J'importe la feuille 'Admin'"
    
    'Save the shared data folder name
    Dim saveDataPath As String
    saveDataPath = ThisWorkbook.Worksheets("Admin").Range("FolderSharedData").value
    
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
        Workbooks.Open(saveDataPath & Application.PathSeparator & "GCF_BD_Sortie.xlsx")
    
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
    targetWorkbook.Sheets(targetSheetName).delete

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

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:ChartOfAccount_Import_All()")
    
    Application.StatusBar = "J'importe le plan comptable"
    
    'Clear all cells, but the headers, in the target worksheet
    wshAdmin.Range("T10").CurrentRegion.Offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entr�e.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx"
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
'    wshBD_Clients.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Call Dynamic_Range_Redefine_Plan_Comptable
        
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    
    Call Output_Timer_Results("modImport:ChartOfAccount_Import_All()", timerStart)

End Sub

Sub Client_List_Import_All() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:Client_List_Import_All()")
    
    Application.StatusBar = "J'importe la liste des clients"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBD_Clients.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_BD_Entr�e.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx" '2024-02-14 @ 07:04
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
    
'    MsgBox _
'        Prompt:="J'ai import� un total de " & _
'            Format(wshBD_Clients.Range("A1").CurrentRegion.Rows.count - 1, _
'            "## ##0") & " clients", _
'        Title:="V�rification du nombre de clients", _
'        Buttons:=vbInformation

    Application.StatusBar = ""

    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    
    Call Output_Timer_Results("modImport:Client_List_Import_All()", timerStart)
        
End Sub

Sub DEB_Recurrent_Import_All() '2024-07-08 @ 08:43
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:DEB_Recurrent_Import_All()")
    
    Application.StatusBar = "J'importe les transactions r�currentes de d�bours�s"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshDEB_Recurrent.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-13 @ 15:09
    sourceTab = "DEB_Recurrent"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshDEB_Recurrent.Range("A1")

    'Copy data, using Range to Range, then close the Master file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False
    
    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshDEB_Recurrent.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshDEB_Recurrent, rng, 1)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set rng = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:DEB_Recurrent_Import_All()", timerStart)

End Sub

Sub DEB_Trans_Import_All() '2024-06-26 @ 18:51
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:DEB_Trans_Import_All()")
    
    Application.StatusBar = "J'importe les transactions de d�bours�s"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshDEB_Trans.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-13 @ 15:09
    sourceTab = "DEB_Trans"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshDEB_Trans.Range("A1")

    'Copy data, using Range to Range, then close the Master file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wshDEB_Trans.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshDEB_Trans, rng, 1)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:DEB_Trans_Import_All()", timerStart)

End Sub

Sub ENC_D�tails_Import_All() '2024-03-07 @ 17:38
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:ENC_D�tails_Import_All()")
    
    Application.StatusBar = "J'importe le d�tail des Encaissements"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_D�tails.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "ENC_D�tails"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshENC_D�tails.Range("A1")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshENC_D�tails.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshENC_D�tails, rng, 1)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:ENC_D�tails_Import_All()", timerStart)

End Sub

Sub ENC_Ent�te_Import_All() '2024-03-07 @ 17:38
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:ENC_Ent�te_Import_All()")
    
    Application.StatusBar = "J'importe le d�tail des Encaissements"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshENC_Ent�te.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "ENC_Ent�te"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshENC_Ent�te.Range("A1")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:36
    Dim rng As Range: Set rng = wshENC_Ent�te.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshENC_Ent�te, rng, 1)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:ENC_Ent�te_Import_All()", timerStart)

End Sub

Sub FAC_Comptes_Clients_Import_All() '2024-03-11 @ 11:33
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:FAC_Comptes_Clients_Import_All()")
    
    Application.StatusBar = "J'importe les comptes clients"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshCAR.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import Comptes_Clients from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "FAC_Comptes_Clients"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshCAR.Range("A2")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshCAR.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_D�tails, rng, 2)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:FAC_Comptes_Clients_Import_All()", timerStart)

End Sub

Sub FAC_D�tails_Import_All() '2024-03-07 @ 17:38
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:FAC_D�tails_Import_All()")
    
    Application.StatusBar = "J'importe le d�tail des Factures"
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_D�tails.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "FAC_D�tails"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshFAC_D�tails.Range("A2")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

   'Setup the format of the worksheet - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wshFAC_D�tails.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_D�tails, rng, 2)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set rng = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:FAC_D�tails_Import_All()", timerStart)

End Sub

Sub FAC_Ent�te_Import_All() '2024-07-11 @ 09:21
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:FAC_Ent�te_Import_All()")
    
    Application.StatusBar = "J'importe les ent�tes de Facture"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Ent�te.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "FAC_Ent�te"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshFAC_Ent�te.Range("A2")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Ent�te.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Ent�te, rng, 2)
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:FAC_Ent�te_Import_All()", timerStart)

End Sub

Sub FAC_Projets_D�tails_Import_All() '2024-07-20 @ 13:25
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:FAC_Projets_D�tails_Import_All()")
    
    Application.StatusBar = "J'importe le d�tail des Projets de Factures"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Projets_D�tails.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "FAC_Projets_D�tails"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshFAC_Projets_D�tails.Range("A1")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wshFAC_Projets_D�tails.Range("A99999").End(xlUp).row
    
    'Delete the rows that column (isD�truite) is set to TRUE
    Dim i As Long
    For i = 2 To lastRow
        If wshFAC_Projets_D�tails.Cells(i, 9) = "VRAI" Then
            wshFAC_Projets_D�tails.rows(i).delete
        End If
    Next i
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wshFAC_Projets_D�tails.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Projets_D�tails, rng, 1)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-20 @ 13:30
    Set destinationRange = Nothing
    Set rng = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:FAC_Projets_D�tails_Import_All()", timerStart)

End Sub

Sub FAC_Projets_Ent�te_Import_All() '2024-07-11 @ 09:21
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:FAC_Projets_Ent�te_Import_All()")
    
    Application.StatusBar = "J'importe les ent�tes de Facture"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Projets_Ent�te.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "FAC_Projets_Ent�te"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshFAC_Projets_Ent�te.Range("A1")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wshFAC_Projets_Ent�te.Range("A99999").End(xlUp).row
    
    'Delete the rows that column (isD�truite) is set to TRUE
    Dim i As Long
    For i = 2 To lastRow
        If wshFAC_Projets_Ent�te.Cells(i, 26) = "VRAI" Then
            wshFAC_Projets_Ent�te.rows(i).delete
        End If
    Next i
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    Dim rng As Range: Set rng = wshFAC_Projets_Ent�te.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshFAC_Projets_Ent�te, rng, 1)
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:FAC_Projets_Ent�te_Import_All()", timerStart)

End Sub

Sub Fournisseur_List_Import_All() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:Fournisseur_List_Import_All()")
    
    Application.StatusBar = "J'importe la liste des fournisseurs"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBD_Fournisseurs.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entr�e.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx" '2024-02-14 @ 07:04
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
    
    'Copy to wshBD_Clients workbook
    wshBD_Fournisseurs.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    Dim rng As Range: Set rng = wshBD_Fournisseurs.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshBD_Fournisseurs, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
'    MsgBox _
'        Prompt:="J'ai import� un total de " & _
'            Format(wshBD_Fournisseurs.Range("A1").CurrentRegion.rows.count - 1, _
'            "##,##0") & " fournisseurs", _
'        Title:="V�rification du nombre de fournisseurs", _
'        Buttons:=vbInformation

    Application.StatusBar = ""

    'Cleaning memory - 2024-07-03 @ 15:45
    Set connStr = Nothing
    Set recSet = Nothing
    
    Call Output_Timer_Results("modImport:Fournisseur_List_Import_All()", timerStart)
        
End Sub

Sub GL_EJ_Auto_Import_All() '2024-03-03 @ 11:36

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:GL_EJ_Auto_Import_All()")
    
    Application.StatusBar = "J'importe les �critures de journal r�currentes"
    
    Application.ScreenUpdating = False
    
    Dim lastUsedRow1 As Long
    lastUsedRow1 = wshGL_EJ_Recurrente.Range("C999").End(xlUp).row
    
    'Clear all cells, but the headers and Columns A & B, in the target worksheet
    If lastUsedRow1 > 1 Then
        wshGL_EJ_Recurrente.Range("C2:I" & lastUsedRow1).ClearContents
    End If
    
    'Import EJ_Auto from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-13 @ 15:09
    sourceTab = "GL_EJ_Auto"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshGL_EJ_Recurrente.Range("C1")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wshGL_EJ_Recurrente.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshGL_EJ_Recurrente, rng, 1)
    
    Call GL_EJ_Auto_Build_Summary '2024-03-14 @ 07:38
    
Clean_Exit:
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:GL_EJ_Auto_Import_All()", timerStart)

End Sub

Sub GL_Trans_Import_All() '2024-03-03 @ 10:13
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:GL_Trans_Import_All()")
    
    Application.StatusBar = "J'importe les transactions du Grand-Livre"
    
    Application.ScreenUpdating = False
    
    Dim saveLastRow As Long
    saveLastRow = wshGL_Trans.Range("A99999").End(xlUp).row
    
    'Clear all cells, but the headers, in the target worksheet
    wshGL_Trans.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-13 @ 15:09
    sourceTab = "GL_Trans"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshGL_Trans.Range("A1")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wshGL_Trans.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshGL_Trans, rng, 1)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:GL_Trans_Import_All()", timerStart)

End Sub

Sub TEC_Import_All() '2024-02-14 @ 06:19
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:TEC_Import_All()")
    
    Application.StatusBar = "J'importe tous les TEC"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshTEC_Local.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-14 @ 06:22
    sourceTab = "TEC"
    
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange
    Dim destinationRange As Range: Set destinationRange = wshTEC_Local.Range("A2")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wshTEC_Local.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshTEC_Local, rng, 2)
    
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing

    Call Output_Timer_Results("modImport:TEC_Import_All()", timerStart)
    
End Sub

