Attribute VB_Name = "Module1"
Option Explicit

Sub FilterProfDate()

    'MsgBox "Temp - Sub FilterProfDate() - Module1"
    'MsgBox "Prof = " & frmSaisieHeures.cmbProfessionnel.Value & " - Date = " & frmSaisieHeures.txtDate.Value
    If Trim(frmSaisieHeures.cmbProfessionnel.value) = "" Or _
        Trim(frmSaisieHeures.txtDate.value) = "" Then
        Exit Sub
    End If
    
    'Base worksheet 'Heures'
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Heures")
    sh.AutoFilterMode = False
    
    'Filtered worksheet 'Heures_Work'
    Dim shFiltered As Worksheet
    Set shFiltered = ThisWorkbook.Sheets("HeuresFiltered")
    shFiltered.UsedRange.Clear
    shFiltered.Activate
    'MsgBox "Temp - La feuille HeuresFiltered devrait être vide ?"
    sh.Activate
        
    sh.UsedRange.AutoFilter 2, frmSaisieHeures.cmbProfessionnel.value
    sh.UsedRange.AutoFilter 3, frmSaisieHeures.txtDate.value
    sh.UsedRange.Select
    'MsgBox "Temp - La feuille Heures ne devrait contenir que les enregistrements filtrés"
    sh.UsedRange.Copy shFiltered.Range("A1")
    shFiltered.Activate
    'MsgBox "Temp - La feuille HeuresFiltered ne devrait contenir que les enregistrements filtrés dans Heures"
    sh.Activate
    sh.AutoFilterMode = False
    sh.ShowAllData
    'MsgBox "Temp - La feuille Heures ne devrait plus avoir de filtres"

End Sub

Sub ImportClientsList()

    'Delete all cells, but the headers in the destination worksheet
    shImportedClients.Range("A1").CurrentRegion.Offset(1, 0).Clear
    
    'Source workbook (closed Excel file) - MUST BE IN THE SAME DIRECTORY
    Dim sourceWorkbook, sourceWorksheet As String
    sourceWorkbook = ThisWorkbook.Path & Application.PathSeparator & _
                     "GCF_Clients.xlsx"
    sourceWorksheet = "Clients"
    
    'ADODB connection
    Dim connStr As ADODB.Connection
    Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = _
        "Provider = Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source = " & sourceWorkbook & ";" & _
        "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset
    Set recSet = New ADODB.Recordset
    
    recSet.ActiveConnection = connStr
    recSet.Source = "SELECT Nom FROM [" & sourceWorksheet & "$]"
        
    recSet.Open
    
    'Copy to destination workbook (actual) into the 'Top2000' worksheet
    shImportedClients.Range("A2").CopyFromRecordset recSet
    
    shImportedClients.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
End Sub
