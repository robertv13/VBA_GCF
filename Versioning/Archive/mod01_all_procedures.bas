Attribute VB_Name = "mod01"
Option Explicit

Sub ImportClientsList() '---------------------------- 'v1.0 - 2023-03-23 @ 07:40

    'Delete all cells, but the headers in the current worksheet
    shImportedClients.Range("A1").CurrentRegion.Offset(1, 0).Clear
    
    'Source workbook (closed Excel file) - MUST BE IN THE SAME DIRECTORY
    Dim sourceWorkbook As String
    Dim sourceWorksheet As String
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
    recSet.Source = "SELECT * FROM [" & sourceWorksheet & "$]"
    recSet.Open
    
    'Copy to destination workbook (actual) into the 'Top2000' worksheet
    shImportedClients.Range("A2").CopyFromRecordset recSet
    shImportedClients.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    MsgBox _
        Prompt:="J'ai un total de " & _
        Format(shImportedClients.Range("A1").CurrentRegion.Rows.count - 1, _
        "## ##0") & " clients", _
        Title:="Vérification du nombre de clients", _
        Buttons:=vbInformation
    
End Sub