Attribute VB_Name = "modImportOusideWorksheet"
Option Explicit

Private Sub ImportOusideWorksheet()

    Dim connection As New ADODB.connection
    
    connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & _
                    ";Extended Properties=""Excel 12.0;HDR=Yes;"";"

    Dim sourceFile As String
    sourceFile = ThisWorkbook.Path & Application.PathSeparator & "Company Sales Data.xlsx"
    
    Dim sourceSheet As String
    sourceSheet = "[Excel 12.0;HDR=YES;DATABASE=" & sourceFile & "]"
    
    Dim query As String
    query = "Insert into [CompanyOut$] Select * From " & sourceSheet & ".[Sales$]"

    connection.Execute query
    
    connection.Close
    
End Sub

    'Instantiate an connection from ADODB
    'Dim connection As New ADODB.connection
    
    'Connection String specific to EXCEL
    'connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & _
                    ";Extended Properties=""Excel 12.0;HDR=Yes;"";"
    
    'Dim sourceFile As String
    'sourceFile = ThisWorkbook.Path & Application.PathSeparator & "Top_2000_companies.xlsx"
    'Debug.Print sourceFile
    
    'Dim sourceSheet As String
    'sourceSheet = "[Excel 12.0;HDR=YES;DATABASE=" & sourceFile & "]"
    
    'Prepare the query from the worksheet (ClientList/shClientList)
    'Dim query As String
    'query = "Insert Into [ClientList$] Select * From " & sourceSheet & ".[Top2000$]"
    'Debug.Print query
    
    'Execute query with connection.Execute
    'connection.Execute query
        
    'connection.Close
    



