Sub ImporterDebRecurrent() '2024-07-08 @ 08:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterDebRecurrent", "", 0)
    
    Application.ScreenUpdating = False
    
    '1. Vider la table locale
    Dim strFeuille As String
    strFeuille = "DEB_Récurrent"
    Dim strTable As String
    strTable = "l_tbl_DEB_Recurrent"
    Call ViderTableau(strFeuille, strTable)
    
    '2. Importer les enregistrements de GCF_MASTER.xlsx
    Dim ws As Worksheet
    Set ws = wsdDEB_Récurrent
    
    'Import GL_Trans from 'GCF_DB_Sortie.xlsx', in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx" '2024-02-13 @ 15:09
    sourceTab = strFeuille & "$"
                     
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
    
    'Copier le recSet vers ws
    If recSet.EOF = False Then
        ws.Range("A2").CopyFromRecordset recSet
    End If
    
    Call AppliquerStyleTable(ws, strTable)
    
    Call DEB_Recurrent_Build_Summary '2025-01-15
    
    'Close resources
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterDebRecurrent", "", startTime)

End Sub
