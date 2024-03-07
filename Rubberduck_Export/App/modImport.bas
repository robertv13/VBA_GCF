Attribute VB_Name = "modImport"
Option Explicit

Sub Client_List_Import_All() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim timerStart As Double: timerStart = Timer
    
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

    Call Output_Timer_Results("Client_List_Import_All()", timerStart)
        
End Sub

Sub TEC_Import_All() '2024-02-14 @ 06:19
    
    Dim timerStart As Double: timerStart = Timer
    
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

    Call Output_Timer_Results("TEC_Import_All()", timerStart)
    
End Sub

Sub ChartOfAccount_Import_All() '2024-02-17 @ 07:21

    Dim timerStart As Double: timerStart = Timer
    
    'Clear all cells, but the headers, in the target worksheet
    wshAdmin.Range("T10").CurrentRegion.Offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx"
    sourceWorksheet = "PlanComptable"
    
    'ADODB connection
    Dim connStr As ADODB.Connection
    Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset
    Set recSet = New ADODB.Recordset
    
    recSet.ActiveConnection = connStr
    recSet.source = "SELECT * FROM [" & sourceWorksheet & "$]"
    recSet.Open
    
    'Copy to wshAdmin workbook
    wshAdmin.Range("T11").CopyFromRecordset recSet
'    wshClientDB.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Call Dynamic_Range_Redefine_Plan_Comptable
        
    Call Output_Timer_Results("ChartOfAccount_Import_All()", timerStart)

End Sub

Sub GL_Trans_Import_All() '2024-03-03 @ 10:13
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    Dim saveLastRow As Long
    saveLastRow = wshGL_Trans.Range("A99999").End(xlUp).row
    
    'Clear all cells, but the headers, in the target worksheet
    wshGL_Trans.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GLTrans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-13 @ 15:09
    sourceTab = "GL_Trans"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range
    Set destinationRange = wshGL_Trans.Range("A1")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    wshGL_Trans.Range("A1").CurrentRegion.EntireColumn.AutoFit
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wshGL_Trans.Range("A999999").End(xlUp).row
    
    'Adjust Formats for all new rows
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
    
'    For r = 2 To lastRow 'RMV - 2024-01-05
'        With wshGL_Trans.Range("A" & r & ":J" & r) 'No_EJ & No.Ligne
'            .Font.ThemeColor = xlThemeColorLight1
'            .Font.TintAndShade = -4.99893185216834E-02
'            .Interior.Pattern = xlNone
'            .Interior.TintAndShade = 0
'            .Interior.PatternTintAndShade = 0
'        End With
'        wshGL_Trans.Range("J" & r).formula = "=ROW()"
'    Next r

    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("GL_Trans_Import_All()", timerStart)

End Sub

Sub GL_JE_Auto_Import_All() '2024-03-03 @ 11:36

    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    Dim saveLastRow As Long
    saveLastRow = wshEJRecurrente.Range("C999").End(xlUp).row
    
    'Clear all cells, but the headers and Columns A & B, in the target worksheet
    wshEJRecurrente.Range("C2:J" & saveLastRow).ClearContents

    'Import GLTrans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-13 @ 15:09
    sourceTab = "EJ_Auto"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange

    Dim destinationRange As Range
    Set destinationRange = wshEJRecurrente.Range("C1")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    wshEJRecurrente.Range("C1").CurrentRegion.Offset(0, 2).EntireColumn.AutoFit
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastUsedRow As Long
    lastUsedRow = wshEJRecurrente.Range("C999").End(xlUp).row
    
    'Adjust Formats for all new rows
    With wshEJRecurrente
        Union(.Range("C2:C" & lastUsedRow), _
            .Range("E2:E" & lastUsedRow)).HorizontalAlignment = xlCenter
        Union(.Range("D2:D" & lastUsedRow), _
            .Range("F2:F" & lastUsedRow)).HorizontalAlignment = xlLeft
        With .Range("G2:H" & lastUsedRow)
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00 $"
        End With
    End With
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("GL_JE_Auto_Import_All()", timerStart)

End Sub

