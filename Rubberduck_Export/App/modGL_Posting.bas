Attribute VB_Name = "modGL_Posting"
Option Explicit

Sub Encaissement_GL_Posting(no As String, dt As Date, nom As String, typeE As String, montant As Currency, desc As String) 'Write/Update to GCF_BD_Sortie.xlsx / GLTrans
    
    Application.ScreenUpdating = False
    
    Dim FullFileName As String, SheetName As String
    FullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    SheetName = "GLTrans"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(No_EJ) AS MaxEJNo FROM [" & SheetName & "$]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim maxEJNo As Long, lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new ID
    Dim nextJENo As Long
    nextJENo = lastJE + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & SheetName & "$] WHERE 1=0", conn, 2, 3
    
    'Debit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("No_EJ").value = nextJENo
        rs.Fields("Date").value = CDate(dt)
        rs.Fields("Num�ro �criture").value = nextJENo
        rs.Fields("Source").value = "Encaissement # " & no
        rs.Fields("Description").value = nom
        rs.Fields("No_Compte").value = "1000" 'Hardcoded
        rs.Fields("Compte").value = "Encaisse" 'Hardcoded
        rs.Fields("D�bit").value = montant
        rs.Fields("Cr�dit").value = 0
        rs.Fields("AutreRemarque").value = desc
    rs.Update
    
    'Credit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("No_EJ").value = nextJENo
        rs.Fields("Date").value = CDate(dt)
        rs.Fields("Num�ro �criture").value = nextJENo
        rs.Fields("Source").value = "Encaissement # " & no
        rs.Fields("Description").value = nom
        rs.Fields("No_Compte").value = "1100" 'Hardcoded
        rs.Fields("Compte").value = "Comptes-Clients" 'Hardcoded
        rs.Fields("D�bit").value = 0
        rs.Fields("Cr�dit").value = montant
        rs.Fields("AutreRemarque").value = desc
    rs.Update

    'Separation line at the end of the Entry
'    rs.AddNew
'        rs.Fields("No_EJ").value = nextJENo
'        rs.Fields("Date").value = CDate(dt)
'        rs.Fields("Num�ro �criture").value = nextJENo
'        rs.Fields("Source").value = "Encaissement"
'        rs.Fields("Description").value = nom
'    rs.Update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

