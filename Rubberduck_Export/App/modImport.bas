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

    'Import Accounts List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
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

    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
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
    wshBD_Clients.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    'Apply standard conditional formatting - 2024-07-08 @ 08:39
    Dim rng As Range: Set rng = wshBD_Clients.Range("A1").CurrentRegion
    Call Apply_Conditional_Formatting(rng, 1)
    
    Application.ScreenUpdating = True
    
'    MsgBox _
'        Prompt:="J'ai importé un total de " & _
'            Format(wshBD_Clients.Range("A1").CurrentRegion.Rows.count - 1, _
'            "## ##0") & " clients", _
'        Title:="Vérification du nombre de clients", _
'        Buttons:=vbInformation

    Application.StatusBar = ""

    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    
    Call Output_Timer_Results("modImport:Client_List_Import_All()", timerStart)
        
End Sub

Sub DEB_Recurrent_Import_All() '2024-07-08 @ 08:43
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:DEB_Recurrent_Import_All()")
    
    Application.StatusBar = "J'importe les transactions récurrentes de déboursés"
    
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

    Dim lastUsedRow As Long
    lastUsedRow = wshDEB_Recurrent.Range("A999999").End(xlUp).row
    
    'Adjust Formats for all new rows
    With wshDEB_Recurrent
        .Range("A2:M" & lastUsedRow).HorizontalAlignment = xlCenter
        .Range("B2:B" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
        .Range("C2:C" & lastUsedRow & _
             ", D2:D" & lastUsedRow & _
             ", E2:E" & lastUsedRow & _
             ", G2:G" & lastUsedRow).HorizontalAlignment = xlLeft
        With .Range("I2:N" & lastUsedRow)
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00 $"
        End With
        .Range("A1").CurrentRegion.EntireColumn.AutoFit
    End With

    'Apply standard conditional formatting - 2024-07-08 @ 08:39
    Dim rng As Range: Set rng = wshDEB_Recurrent.Range("A1").CurrentRegion
    Call Apply_Conditional_Formatting(rng, 2)
    
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
    
    Application.StatusBar = "J'importe les transactions de déboursés"
    
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

    Dim lastUsedRow As Long
    lastUsedRow = wshDEB_Trans.Range("A999999").End(xlUp).row
    
    'Adjust Formats for all new rows
    With wshDEB_Trans
        .Range("A2:P" & lastUsedRow).HorizontalAlignment = xlCenter
        .Range("B2:B" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
        .Range("C2:C" & lastUsedRow & _
             ", D2:D" & lastUsedRow & _
             ", F2:F" & lastUsedRow & _
             ", H2:H" & lastUsedRow & _
             ", O2:O" & lastUsedRow).HorizontalAlignment = xlLeft
        With .Range("J2:N" & lastUsedRow)
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00 $"
        End With
        .Range("A1").CurrentRegion.EntireColumn.AutoFit
    End With
    
    'Apply standard conditional formatting - 2024-07-08 @ 08:39
    Dim rng As Range: Set rng = wshDEB_Trans.Range("A1").CurrentRegion
    Call Apply_Conditional_Formatting(rng, 2)
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:DEB_Trans_Import_All()", timerStart)

End Sub

Sub FAC_Comptes_Clients_Import_All() '2024-03-11 @ 11:33
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:FAC_Comptes_Clients_Import_All()")
    
    Application.StatusBar = "J'importe les comptes clients"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import Comptes_Clients from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "FAC_Comptes_Clients"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshFAC.Range("A2")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wshFAC.Range("A99999").End(xlUp).row
    
    'Adjust Formats for all new rows
    With wshFAC
        .Range("A3:B" & lastRow & ", D3:F" & lastRow & ", J3:J" & lastRow).HorizontalAlignment = xlCenter
        .Range("C3:C" & lastRow).HorizontalAlignment = xlLeft
        .Range("G3:I" & lastRow).HorizontalAlignment = xlRight
        .Range("B3:B" & lastRow).NumberFormat = "dd/mm/yyyy"
        .Range("G3:I" & lastRow).NumberFormat = "#,##0.00 $"
        .Range("A1").CurrentRegion.EntireColumn.AutoFit
    End With

    'Apply standard conditional formatting - 2024-07-08 @ 08:39
    Dim rng As Range: Set rng = wshFAC.Range("A1").CurrentRegion
    Call Apply_Conditional_Formatting(rng, 2)
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:FAC_Comptes_Clients_Import_All()", timerStart)

End Sub

Sub FAC_Détails_Import_All() '2024-03-07 @ 17:38
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:FAC_Détails_Import_All()")
    
    Application.StatusBar = "J'importe le détail des Factures"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Détails.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "FAC_Détails"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshFAC_Détails.Range("A2")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    wshFAC_Détails.Range("A1").CurrentRegion.EntireColumn.AutoFit
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wshFAC_Détails.Range("A99999").End(xlUp).row
    
    'Adjust Formats for all rows
    With wshFAC_Détails
        .Range("A4:A" & lastRow & ", C4:C" & lastRow & ", F4:F" & lastRow & ", G4:G" & lastRow).HorizontalAlignment = xlCenter
        .Range("B4:B" & lastRow).HorizontalAlignment = xlLeft
        .Range("D4:E" & lastRow).HorizontalAlignment = xlRight
        .Range("C4:C" & lastRow).NumberFormat = "#,##0.00"
        .Range("D4:E" & lastRow).NumberFormat = "#,##0.00 $"
        .Range("H4:H" & lastRow & ",J4:J" & lastRow & ",L4:L" & lastRow & ",N4:T" & lastRow).NumberFormat = "#,##0.00 $"
        .Range("O4:O" & lastRow & ",Q4:Q" & lastRow).NumberFormat = "#0.000 %"
    End With

    'Apply standard conditional formatting - 2024-07-08 @ 08:39
    Dim rng As Range: Set rng = wshFAC_Détails.Range("A1").CurrentRegion
    Call Apply_Conditional_Formatting(rng, 2)
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:FAC_Détails_Import_All()", timerStart)

End Sub

Sub FAC_Entête_Import_All() '2024-03-13 @ 09:56
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:FAC_Entête_Import_All()")
    
    Application.StatusBar = "J'importe les entêtes de Facture"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshFAC_Entête.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    sourceTab = "FAC_Entête"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range: Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range: Set destinationRange = wshFAC_Entête.Range("A2")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    wshFAC_Entête.Range("A1").CurrentRegion.EntireColumn.AutoFit
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wshFAC_Entête.Range("A99999").End(xlUp).row
    
    'Adjust Formats for all rows
    With wshFAC_Entête
        .Range("A3:C" & lastRow).HorizontalAlignment = xlCenter
        .Range("B3:B" & lastRow).NumberFormat = "dd/mm/yyyy"
        .Range("D3:H" & lastRow & ",J3:J" & lastRow & ",L3:L" & lastRow & ",N3:N" & lastRow).HorizontalAlignment = xlLeft
        .Range("I3:I" & lastRow & ",K3:K" & lastRow & ",M3:M" & lastRow & ",O3:U" & lastRow).HorizontalAlignment = xlRight
        .Range("I3:I" & lastRow & ",K3:K" & lastRow & ",M3:M" & lastRow & ",O3:U" & lastRow).NumberFormat = "#,##0.00 $"
        .Range("P3:P" & lastRow & ",R3:R" & lastRow).NumberFormat = "#0.000 %"
    End With

    'Apply standard conditional formatting - 2024-07-08 @ 08:39
    Dim rng As Range: Set rng = wshFAC_Entête.Range("A1").CurrentRegion
    Call Apply_Conditional_Formatting(rng, 2)
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing
    
    Call Output_Timer_Results("modImport:FAC_Entête_Import_All()", timerStart)

End Sub

Sub Fournisseur_List_Import_All() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:Fournisseur_List_Import_All()")
    
    Application.StatusBar = "J'importe la liste des fournisseurs"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBD_Fournisseurs.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
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
    
    'Copy to wshBD_Clients workbook
    wshBD_Fournisseurs.Range("A2").CopyFromRecordset recSet
    wshBD_Fournisseurs.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    'Apply standard conditional formatting - 2024-07-08 @ 08:39
    Dim rng As Range: Set rng = wshBD_Fournisseurs.Range("A1").CurrentRegion
    Call Apply_Conditional_Formatting(rng, 1)
    
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
    
    Call Output_Timer_Results("modImport:Fournisseur_List_Import_All()", timerStart)
        
End Sub

Sub GL_EJ_Auto_Import_All() '2024-03-03 @ 11:36

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modImport:GL_EJ_Auto_Import_All()")
    
    Application.StatusBar = "J'importe les écritures de journal récurrentes"
    
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
    wshGL_EJ_Recurrente.Range("C1").CurrentRegion.Offset(0, 2).EntireColumn.AutoFit
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Get the last used rows AFTER the copy
    lastUsedRow1 = wshGL_EJ_Recurrente.Range("C999").End(xlUp).row
    
    'Adjust Formats for all rows
    With wshGL_EJ_Recurrente
        Union(.Range("C2:C" & lastUsedRow1), _
              .Range("E2:E" & lastUsedRow1)).HorizontalAlignment = xlCenter
        Union(.Range("D2:D" & lastUsedRow1), _
              .Range("F2:F" & lastUsedRow1), _
              .Range("I2:I" & lastUsedRow1)).HorizontalAlignment = xlLeft
        With .Range("G2:H" & lastUsedRow1)
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00 $"
        End With
    End With
    
    Call GL_EJ_Auto_Build_Summary '2024-03-14 @ 07:38
    
Clean_Exit:
    'Apply standard conditional formatting - 2024-07-08 @ 08:39
    Dim rng As Range: Set rng = wshGL_EJ_Recurrente.Range("A1").CurrentRegion
    Call Apply_Conditional_Formatting(rng, 1)
    
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
    wshGL_Trans.Range("A1").CurrentRegion.EntireColumn.AutoFit
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wshGL_Trans.Range("A999999").End(xlUp).row
    
    'Adjust Formats for all the rows
    With wshGL_Trans
        .Range("A" & 2 & ":J" & lastRow).HorizontalAlignment = xlCenter
        .Range("B" & 2 & ":B" & lastRow).NumberFormat = "dd/mm/yyyy"
        .Range("C" & 2 & ":C" & lastRow & _
            ", D" & 2 & ":D" & lastRow & _
            ", F" & 2 & ":F" & lastRow & _
            ", I" & 2 & ":I" & lastRow) _
                .HorizontalAlignment = xlLeft
        With .Range("G" & 2 & ":H" & lastRow)
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00 $"
        End With
        With .Range("A" & 2 & ":A" & lastRow) _
            .Range("J" & 2 & ":J" & lastRow).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    
    Dim firstRowJE As Long, lastRowJE As Long
    Dim r As Long
    
    'Apply standard conditional formatting - 2024-07-08 @ 08:39
    Dim rng As Range: Set rng = wshGL_Trans.Range("A1").CurrentRegion
    Call Apply_Conditional_Formatting(rng, 1)
    
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
    wshTEC_Local.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Arrange formats on all rows
    Dim lastRow As Long
    lastRow = wshTEC_Local.Range("A99999").End(xlUp).row
    
    With wshTEC_Local
        .Range("A3" & ":P" & lastRow).HorizontalAlignment = xlCenter
        With .Range("F3:F" & lastRow & ",G3:G" & lastRow & ",I3:I" & lastRow & ",O3:O" & lastRow)
            .HorizontalAlignment = xlLeft
        End With
        .Range("H3:H" & lastRow).NumberFormat = "#0.00"
        .Range("K3:K" & lastRow).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End With
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set destinationRange = Nothing
    Set sourceRange = Nothing

    Call Output_Timer_Results("modImport:TEC_Import_All()", timerStart)
    
End Sub

