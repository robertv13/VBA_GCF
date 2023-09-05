Attribute VB_Name = "Module1"
Option Explicit

Sub GetDataFromMovies()

    'Delete all cells, but the headers in the destination worksheet
    Sheet1.Range("A1").CurrentRegion.Offset(1, 0).Clear
    
    'Source workbook (closed Excel file) - MUST BE IN THE SAME DIRECTORY
    Dim sourceWorkbook As String
    sourceWorkbook = ThisWorkbook.Path & Application.PathSeparator & _
        "Movies.xlsx"
    
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
    recSet.Source = "SELECT * FROM [Sheet1$]"
        
    recSet.Open
    
    'Copy to destination workbook (actual) into the 'sheet1' worksheet
    Sheet1.Range("A2").CopyFromRecordset recSet
    
    Sheet1.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
End Sub
