Sub ImporterCCRegularisations() '2025-01-05 @ 11:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterCCRegularisations", "", 0)
    
    Application.ScreenUpdating = False
    
    '1. Vider la table locale
    Dim strFeuille As String
    strFeuille = "CC_Régularisations"
    Dim strTable As String
    strTable = "l_tbl_CC_Régularisations"
    Call ViderTableau(strFeuille, strTable)
    
    '2. Importer les enregistrements de GCF_MASTER.xlsx
    Dim ws As Worksheet
    Set ws = wsdCC_Régularisations
    
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

'   'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:35
'    Dim rng As Range: Set rng = wsdCC_Régularisations.Range("A1").CurrentRegion
'    Call AppliquerFormatColonnesParTable(wsdCC_Régularisations, rng, 1)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterCCRegularisations", "", startTime)

End Sub