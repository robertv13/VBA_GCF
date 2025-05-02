Attribute VB_Name = "modImport"
Option Explicit

Sub ImporterPlanComptable() '2024-02-17 @ 07:21

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterPlanComptable", "", 0)

    'Clear all cells, but the headers, in the target worksheet
    wsdADMIN.Range("T10").CurrentRegion.offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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

    'Copy to wsdAdmin workbook
    wsdADMIN.Range("T11").CopyFromRecordset recSet

    'Close resource
    recSet.Close
    connStr.Close

    Call RedefinirDnrPlanComptable

    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing

    Call Log_Record("modImport:ImporterPlanComptable", "", startTime)

End Sub

Sub ImporterClients() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterClients", "", 0)
    
    Application.ScreenUpdating = False
    
    'Worksheet recevant les données importées
    Dim ws As Worksheet
    Set ws = wsdBD_Clients
    
    'Efface toutes les lignes, sauf la ligne d'entête
    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents
    
    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copier le recSet vers ws
    ws.Range("A2").CopyFromRecordset recSet
    
    'Redimensionner le tableau & appliquer le format
    Dim tableName As String
    tableName = "l_tbl_BD_Clients"
    
    Call RedimensionnerTable(ws, tableName)
    Call AppliquerStyleTable(ws, tableName)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterClients", "", startTime)

End Sub

Sub ImporterDebRecurrent() '2024-07-08 @ 08:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterDebRecurrent", "", 0)
    
    Application.ScreenUpdating = False
    
    'Feuille qui sera importée
    Dim ws As Worksheet
    Set ws = wsdDEB_Récurrent
    
    Dim strNomTable As String
    strNomTable = "l_tbl_DEB_Recurrent"
    Dim lo As ListObject
    Set lo = ws.ListObjects(strNomTable)
    
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ShowAutoFilter Then
            lo.AutoFilter.ShowAllData
        End If
        lo.DataBodyRange.Delete
    End If
    
    'Import GL_Trans from 'GCF_DB_Sortie.xlsx', in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    'Copy to wsdDEB_Récurrent workbook
    ws.Range("A2").CopyFromRecordset recSet

    'Redimensionner la table pour refléter le nombre exact de lignes occupées
    Dim derLigne As Long
    derLigne = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    Set lo = ws.ListObjects(strNomTable)
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1), ws.Cells(derLigne, recSet.Fields.count))
    
    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wsdDEB_Récurrent.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdDEB_Récurrent, rng, 1)
    
    Call DEB_Recurrent_Build_Summary '2025-01-15 @
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set lo = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterDebRecurrent", "", startTime)

End Sub

Sub ImporterDebTrans() '2024-06-26 @ 18:51
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterDebTrans", "", 0)
    
    Application.ScreenUpdating = False
    
    'Feuille qui sera importée
    Dim ws As Worksheet
    Set ws = wsdDEB_Trans
    
    Dim strNomTable As String
    strNomTable = "l_tbl_DEB_Trans"
    Dim lo As ListObject
    Set lo = ws.ListObjects(strNomTable)
    
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ShowAutoFilter Then
            lo.AutoFilter.ShowAllData
        End If
        lo.DataBodyRange.Delete
    End If
    
    'Import DEB_Trans from 'GCF_BD_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdDEB_Trans workbook after erasing actual lines
    wsdDEB_Trans.Range("A2").CopyFromRecordset recSet
    
    'Redimensionner la table pour refléter le nombre exact de lignes occupées
    Dim derLigne As Long
    derLigne = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    Set lo = ws.ListObjects(strNomTable)
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1), ws.Cells(derLigne, recSet.Fields.count))
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wsdDEB_Trans.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdDEB_Trans, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set lo = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterDebTrans", "", startTime)

End Sub

Sub ImporterEncDetails() '2025-01-16 @ 16:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterEncDetails", "", 0)
    
    Application.ScreenUpdating = False
    
    'Feuille qui sera importée
    Dim ws As Worksheet
    Set ws = wsdENC_Détails
    
    Dim strNomTable As String
    strNomTable = "l_tbl_ENC_Détails"
    Dim lo As ListObject
    Set lo = ws.ListObjects(strNomTable)
    
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ShowAutoFilter Then
            lo.AutoFilter.ShowAllData
        End If
        lo.DataBodyRange.Delete
    End If
    
    'Import ENC_Détails from 'GCF_BD_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdENC_Détails worksheet
    wsdENC_Détails.Range("A2").CopyFromRecordset recSet

    'Redimensionner la table pour refléter le nombre exact de lignes occupées
    Dim derLigne As Long
    derLigne = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    Set lo = ws.ListObjects(strNomTable)
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1), ws.Cells(derLigne, recSet.Fields.count))
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wsdENC_Détails.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdENC_Détails, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set lo = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterEncDetails", "", startTime)

End Sub

Sub ImporterEncEntete() '2025-03-10 @ 17:08
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterEncEntete", "", 0)
    
    Application.ScreenUpdating = False
    
    'Feuille qui sera importée
    Dim ws As Worksheet
    Set ws = wsdENC_Entête
    
    Dim strNomTable As String
    strNomTable = "l_tbl_ENC_Entête"
    Dim lo As ListObject
    Set lo = ws.ListObjects(strNomTable)
    
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ShowAutoFilter Then
            lo.AutoFilter.ShowAllData
        End If
        lo.DataBodyRange.Delete
    End If
    
    'Import ENC_Entête from 'GCF_BD_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdENC_Entête worksheet
    wsdENC_Entête.Range("A2").CopyFromRecordset recSet
    
    'Redimensionner la table pour refléter le nombre exact de lignes occupées
    Dim derLigne As Long
    derLigne = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    Set lo = ws.ListObjects(strNomTable)
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1), ws.Cells(derLigne, recSet.Fields.count))
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:36
    Dim rng As Range: Set rng = wsdENC_Entête.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdENC_Entête, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set lo = Nothing
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterEncEntete", "", startTime)

End Sub

Sub ImporterCCRegularisations() '2025-01-05 @ 11:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterCCRegularisations", "", 0)
    
    Application.ScreenUpdating = False
    
    'Feuille qui sera importée
    Dim ws As Worksheet
    Set ws = wsdCC_Régularisations
    
    Dim strNomTable As String
    strNomTable = "tbl_REGUL_Détails"
    Dim lo As ListObject
    Set lo = ws.ListObjects(strNomTable)
    
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ShowAutoFilter Then
            lo.AutoFilter.ShowAllData
        End If
        lo.DataBodyRange.Delete
    End If
    
    'Import CC_Régularisations from 'GCF_BD_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdCC_Régularisations worksheet
    If recSet.EOF = False Then
        wsdCC_Régularisations.Range("A2").CopyFromRecordset recSet
    End If

    'Redimensionner la table pour refléter le nombre exact de lignes occupées
    Dim derLigne As Long
    derLigne = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    Set lo = ws.ListObjects(strNomTable)
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1), ws.Cells(derLigne, recSet.Fields.count))
   
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wsdCC_Régularisations.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdCC_Régularisations, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set lo = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterCCRegularisations", "", startTime)

End Sub

Sub ImporterFacComptesClients() '2024-08-07 @ 17:41
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacComptesClients", "", 0)
    
    Application.ScreenUpdating = False
    
    'Feuille qui sera importée
    Dim ws As Worksheet
    Set ws = wsdFAC_Comptes_Clients
    
    Dim strNomTable As String
    strNomTable = "tblFAC_Comptes_Clients"
    Dim lo As ListObject
    Set lo = ws.ListObjects(strNomTable)
    
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ShowAutoFilter Then
            lo.AutoFilter.ShowAllData
        End If
        lo.DataBodyRange.Delete
    End If
    
    'Import FAC_Comptes_Clients from 'GCF_DB_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    wsdFAC_Comptes_Clients.Range("A3").CopyFromRecordset recSet
    
    'Redimensionner la table pour refléter le nombre exact de lignes occupées
    Dim derLigne As Long
    derLigne = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    Set lo = ws.ListObjects(strNomTable)
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1), ws.Cells(derLigne, recSet.Fields.count))
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:32
    Dim rng As Range: Set rng = wsdFAC_Comptes_Clients.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdFAC_Comptes_Clients, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set lo = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterFacComptesClients", "", startTime)

End Sub

Sub ImporterFacDetails() '2024-03-07 @ 17:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacDetails", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wsdFAC_Détails.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdFAC_Détails workbook
    wsdFAC_Détails.Range("A3").CopyFromRecordset recSet

   'Setup the format of the worksheet - 2024-07-20 @ 18:35
    Dim rng As Range: Set rng = wsdFAC_Détails.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdFAC_Détails, rng, 2)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ImporterFacDetails", "", startTime)

End Sub

Sub ImporterFacEntete() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacEntete", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = wsdFAC_Entête
    
    Dim strNomTable As String
    strNomTable = "l_tbl_FAC_Entête"
    Dim lo As ListObject
    Set lo = ws.ListObjects(strNomTable)
    
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ShowAutoFilter Then
            lo.AutoFilter.ShowAllData
        End If
        lo.DataBodyRange.Delete
    End If
    
    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdFAC_Entête workbook
    wsdFAC_Entête.Range("A3").CopyFromRecordset recSet
    
    'Redimensionner la table pour inclure uniquement les nouvelles données
    Dim derLigne As Long
    derLigne = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    Set lo = ws.ListObjects(strNomTable)
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1), ws.Cells(derLigne, recSet.Fields.count))
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wsdFAC_Entête.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdFAC_Entête, rng, 2)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set lo = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterFacEntete", "", startTime)

End Sub

Sub ImporterFacSommaireTaux() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacSommaireTaux", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wsdFAC_Sommaire_Taux.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import FAC_Sommaire_Taux from 'GCF_BD_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdFAC_Entête workbook
    wsdFAC_Sommaire_Taux.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:37
    Dim rng As Range: Set rng = wsdFAC_Sommaire_Taux.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdFAC_Sommaire_Taux, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ImporterFacSommaireTaux", "", startTime)

End Sub

Sub ImporterFacProjetsDetails() '2024-07-20 @ 13:25
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacProjetsDetails", "", 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Détails
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import FAC_Projets_Détails from 'GCF_DB_MASTER.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy all rows to wsdFAC_Projets_Détails workbook
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
        Call AppliquerFormatColonnesParTable(wsdFAC_Projets_Détails, dataRange, 1)
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
    
    Call Log_Record("modImport:ImporterFacProjetsDetails", "", startTime)

End Sub

Sub ImporterFacProjetsEntete() '2024-07-11 @ 09:21
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacProjetsEntete", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Entête
    
    'Clear all cells, but the headers, in the target worksheet
    ws.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdFAC_Projets_Entête workbook
    If recSet.RecordCount > 0 Then
        ws.Range("A2").CopyFromRecordset recSet
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    'Delete the rows that column (isDétruite) is set to TRUE
    Dim i As Long
    If lastRow >= 2 Then
        For i = lastRow To 2 Step -1
            If UCase$(ws.Cells(i, 26).value) = "VRAI" Or _
                ws.Cells(i, 26).value = -1 Then
                ws.Rows(i).Delete
            End If
        Next i
    End If
    
   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 1 Then
        Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
        Call AppliquerFormatColonnesParTable(ws, rng, 1)
    End If
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterFacProjetsEntete", "", startTime)

End Sub

Sub ImporterFournisseurs() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFournisseurs", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wsdBD_Fournisseurs.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdBD_Fournisseurs workbook
    wsdBD_Fournisseurs.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    Dim rng As Range: Set rng = wsdBD_Fournisseurs.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdBD_Fournisseurs, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ImporterFournisseurs", "", startTime)

End Sub

Sub ImporterEJRecurrente() '2024-03-03 @ 11:36

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterEJRecurrente", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim lastUsedRow As Long
    lastUsedRow = wsdGL_EJ_Recurrente.Cells(wsdGL_EJ_Recurrente.Rows.count, "C").End(xlUp).row
    
    'Clear all cells, but the headers and Columns A & B, in the target worksheet
    If lastUsedRow > 1 Then
        wsdGL_EJ_Recurrente.Range("C2:I" & lastUsedRow).ClearContents
    End If
    
    'Import EJ_Auto from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdGL_EJ_Recurrente workbook
    wsdGL_EJ_Recurrente.Range("A2").CopyFromRecordset recSet

   'Setup the format of the worksheet using a Sub
    Dim rng As Range: Set rng = wsdGL_EJ_Recurrente.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdGL_EJ_Recurrente, rng, 1)
    
    Call GL_EJ_Recurrente_Build_Summary '2024-03-14 @ 07:38
    
Clean_Exit:
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ImporterEJRecurrente", "", startTime)

End Sub

Sub ImporterGLTransactions() '2024-03-03 @ 10:13
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterGLTransactions", "", 0)
    
    Application.ScreenUpdating = False
    
    'Worksheet recevant les données importées
    Dim wsLocal As Worksheet: Set wsLocal = wsdGL_Trans
    
    'Effacer toutes les lignes, sauf la ligne d'entête
    Dim saveLastRow As Long
    saveLastRow = wsLocal.Cells(wsLocal.Rows.count, 1).End(xlUp).row
    If saveLastRow > 1 Then
        wsLocal.Range("A1").CurrentRegion.offset(1, 0).ClearContents
    End If

    'Import GL_Trans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    Call AppliquerFormatColonnesParTable(wsLocal, rng, 1)

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set wsLocal = Nothing
    
    Call Log_Record("modImport:ImporterGLTransactions", "", startTime)

End Sub

Sub ImporterTEC()                             '2024-02-14 @ 06:19
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterTEC", "", 0)
    
'    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wsdTEC_Local
    
    'Clear all cells, but the headers, in the destination worksheet
    ws.Range("A1").CurrentRegion.offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
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
    
    'Copy to wsdTEC_Local workbook
    ws.Range("A3").CopyFromRecordset recSet

    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterTEC", "", startTime)

End Sub

Sub RedimensionnerTable(targetSheet As Worksheet, tableName As String) '2025-03-11 @ 07:28

    'Trouver la table
    Dim tbl As ListObject
    Set tbl = targetSheet.ListObjects(tableName)
    
    'Déterminer la dernière ligne et colonne des nouvelles données
    Dim lastRow As Long
    Dim lastCol As Long
    With targetSheet
        lastRow = .Cells(.Rows.count, tbl.Range.Column).End(xlUp).row
        lastCol = .Cells(tbl.Range.row, .Columns.count).End(xlToLeft).Column
    End With
    
    'Redimensionner la tableàsa
    tbl.Resize targetSheet.Range(tbl.Range.Cells(1, 1), targetSheet.Cells(lastRow, lastCol))
    
End Sub

Sub AppliquerStyleTable(targetSheet As Worksheet, tableName As String) '2025-03-11 @ 07:28

    'Identifier la table
    Dim tbl As ListObject
    Set tbl = targetSheet.ListObjects(tableName)
    
    'Appliquer un style existant à la table
    tbl.tableStyle = "TableStyleMedium2"
    
End Sub


