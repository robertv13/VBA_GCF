Attribute VB_Name = "modWriteToClosedExcelFile"
Option Explicit

Sub AddOneRecordToDB()                           '2023-12-05 @ 03:39

    Dim fileLocation As String
    Dim fileToOpen As Workbook
    Dim userName
    
    userName = Environ("USERNAME")
    
    Application.ScreenUpdating = False
    
    fileLocation = "C:\VBA\GC_FISCALITÉ" & Application.PathSeparator & _
                   "DataFiles" & Application.PathSeparator & _
                   "GCF_BD_Test.xlsx"
 
    'Open the file
    Set fileToOpen = Workbooks.Open(fileLocation)
    
    'Code to write something
    Dim rowAvail As Long
    rowAvail = fileToOpen.Worksheets(1).Range("A999999").End(xlUp).Row + 1
    fileToOpen.Worksheets(1).Range("A" & rowAvail).Value = "=ROW()-1"
    fileToOpen.Worksheets(1).Range("B" & rowAvail).Value = Now
    fileToOpen.Worksheets(1).Range("C" & rowAvail).Value = userName
    
    'Save & Close the file
    fileToOpen.Close True
    
    Application.ScreenUpdating = True

End Sub

Sub GetNumberOfRecordsFromDB()

    Dim fileLocation As String
    Dim fileToOpen As Workbook
    
    Application.ScreenUpdating = False
    
    fileLocation = ThisWorkbook.Path & Application.PathSeparator & _
                   "DataFiles" & Application.PathSeparator & _
                   "GCF_DB_Testxlsx"
 
    'Open the closed file
    Set fileToOpen = Workbooks.Open(fileLocation)
    
    'Code to determine how many records are used
    Dim rowLast As Long
    rowLast = fileToOpen.Worksheets(1).Range("A999999").End(xlUp).Row
    wshCode.Range("G6").Value = rowLast - 1
    
    'Close the file without saving any changes
    fileToOpen.Close False
    
    Application.ScreenUpdating = True

End Sub

Sub ChatGPT_AddRecordToEndOfWorksheet()          'Write to a closed -OR- open .xlsx file
    Dim FilePath As String
    Dim sheetName As String
    Dim conn As Object
    Dim rs As Object
    Dim strConn As String
    Dim strSQL As String
    Dim nextID As Long

    Application.ScreenUpdating = False
    
    'Set the file full path and worksheet name, assuming the same directory -OR- underneath
    FilePath = ThisWorkbook.Path & Application.PathSeparator & _
               "DataFiles" & Application.PathSeparator & _
               "GCF_DB_Testxlsx"
    sheetName = "Feuil1"

    'Initialize connection, connection string & open the connection
    Set conn = CreateObject("ADODB.Connection")
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    conn.Open strConn

    'Initialize recordset
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    strSQL = "SELECT MAX(ID) AS MaxID FROM [" & sheetName & "$];"

    'Open the recordset with the select command
    rs.Open strSQL, conn, 2, 3

    'Get the next available ID
    nextID = rs.Fields("MaxID").Value + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
    rs.AddNew
    
    'Add fields to the record, before updating it
    rs.Fields("ID").Value = nextID
    rs.Fields("Timestamp").Value = Format(Now, "dd-mm-yyyy hh:mm:ss")
    
    'Update the recordset (create the record)
    rs.Update

    'Close recordset and connection
    rs.Close
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

Sub AddRecordToEndOfWorksheet(FilePath As String, sheetName As String) 'Write to a closed -OR- open .xlsx file
    Dim conn As Object
    Dim rs As Object
    Dim strConn As String
    Dim strSQL As String
    Dim nextID As Long
    Dim userName As String

    Application.ScreenUpdating = False
    
    userName = Environ("USERNAME")
    
    'Initialize connection, connection string & open the connection
    Set conn = CreateObject("ADODB.Connection")
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    conn.Open strConn

    'Initialize recordset
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    strSQL = "SELECT MAX(ID) AS MaxID FROM [" & sheetName & "$];"

    'Open the recordset with the select command
    rs.Open strSQL, conn, 2, 3

    'Get the next available ID
    nextID = rs.Fields("MaxID").Value + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
    rs.AddNew
    
    'Add fields to the record, before updating it
    rs.Fields("ID").Value = nextID
    rs.Fields("Timestamp").Value = Format(Now, "dd-mm-yyyy hh:mm:ss")
    rs.Fields("userName").Value = userName
    
    'Update the recordset (create the record)
    rs.Update

    'Close recordset and connection
    rs.Close
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

Sub AddXRecordToEndOfWorksheet(FilePath As String, sheetName As String, numberRecords As Long) 'Write to a closed -OR- open .xlsx file
    Dim conn As Object
    Dim rs As Object
    Dim strConn As String
    Dim strSQL As String
    Dim nextID As Long
    Dim userName As String
    
    Application.ScreenUpdating = False
    
    userName = Environ("USERNAME")

    'Initialize connection, connection string & open the connection
    Set conn = CreateObject("ADODB.Connection")
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    conn.Open strConn

    'Initialize recordset
    Set rs = CreateObject("ADODB.Recordset")

    Dim i As Integer
    For i = 1 To numberRecords
        'SQL select command to find the next available ID
        strSQL = "SELECT MAX(ID) AS MaxID FROM [" & sheetName & "$];"
    
        'Open the recordset with the select command
        rs.Open strSQL, conn, 2, 3
    
        'Get the next available ID
        nextID = rs.Fields("MaxID").Value + 1
    
        'Close the previous recordset, no longer needed and open an empty recordset
        rs.Close
        rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
        rs.AddNew
        
        'Add fields to the record, before updating it
        rs.Fields("ID").Value = nextID
        rs.Fields("timeStamp").Value = Format(Now, "dd-mm-yyyy hh:mm:ss")
        rs.Fields("userName").Value = userName
        
        'Update the recordset (create the record)
        rs.Update
        rs.Close
    Next i
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

Sub Write10RecordUsingADO()

    Dim FilePath As String
    Dim sheetName As String

    FilePath = ThisWorkbook.Path & Application.PathSeparator & _
               "DataFiles" & Application.PathSeparator & _
               "GCF_DB_Testxlsx"
    sheetName = "Feuil1"

    Dim i As Integer
    For i = 1 To 10
        AddRecordToEndOfWorksheet FilePath, sheetName
    Next i
    
End Sub

Sub WriteRecordsUsingADO()

    Dim FilePath As String
    Dim sheetName As String

    FilePath = ThisWorkbook.Path & Application.PathSeparator & _
               "DataFiles" & Application.PathSeparator & _
               "GCF_DB_Testxlsx"
    sheetName = "Feuil1"

    'Sub AddXRecordToEndOfWorksheet Nom complet du fichier, Nom de la feuille, Nombre d'enregistrements à créer
    AddXRecordToEndOfWorksheet FilePath, sheetName, 1000
    
End Sub


