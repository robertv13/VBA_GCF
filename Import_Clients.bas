Sub ImporterClients() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterClients", "", 0)
    
    Application.ScreenUpdating = False
    
    '1. Vider la table locale
    Dim strFeuille As String
    strFeuille = "BD_Clients"
    Dim strTable As String
    strTable = "l_tbl_BD_Clients"
    Call ViderTableau(strFeuille, strTable)
    
    '2. Importer les enregistrements de GCF_MASTER.xlsx
    Dim ws As Worksheet
    Set ws = wsdBD_Clients
    
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
        .Open
    End With
    
    'Copier le recSet vers ws
    If recSet.EOF = False Then
        ws.Range("A2").CopyFromRecordset recSet
    End If
    
    Call AppliquerStyleTable(ws, strTable)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterClients", "", startTime)

End Sub
