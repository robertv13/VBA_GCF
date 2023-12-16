Attribute VB_Name = "modTEC"
Option Explicit

Global Const rmv_modeInitial As Integer = 1
Global Const rmv_modeCreation As Integer = 2
Global Const rmv_modeAffichage As Integer = 3
Global Const rmv_modeModification As Integer = 4

Global rmv_state As Integer

Global savedClient As String
Global savedActivite As String
Global savedHeures As String
Global savedFacturable As String
Global savedCommNote As String

Global Const gAppVersion As String = "v1.0.7"

Sub ImportClientList()                                          '---------------- 2023-11-12 @ 07:28
    
    'Clear all cells, but the headers, in the worksheet
    wshClientDB.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_Clients.xlsx. In order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("SharedFolder").value & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx" '2023-12-15 @ 07:23
    sourceWorksheet = "Clients"
    
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
    recSet.Source = "SELECT * FROM [" & sourceWorksheet & "$]"
    recSet.Open
    
    'Copy to wshClientDB workbook
    wshClientDB.Range("A2").CopyFromRecordset recSet
    wshClientDB.Range("A:B").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
'    MsgBox _
'        Prompt:="J'ai importé un total de " & _
'            Format(wshClientDB.Range("A1").CurrentRegion.Rows.count - 1, _
'            "## ##0") & " clients", _
'        Title:="Vérification du nombre de clients", _
'        Buttons:=vbInformation
        
End Sub

Sub ImportTEC()
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the worksheet
    wshBaseHours.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String
    sourceWorkbook = wshAdmin.Range("SharedFolder").value & Application.PathSeparator & "GCF_BD_Sortie.xlsx" '2023-12-15 @ 19:15

    'Set up source and destination ranges
    Dim sourceRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets("TEC").UsedRange
    Debug.Print "Je vais copier toutes les cellules du Range = " & sourceRange.Address & " dans BaseHours!"

    Dim destinationRange As Range
    Set destinationRange = wshBaseHours.Range("A2")

    'Copy data
    sourceRange.Copy destinationRange

    'Close the source workbook
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Autofit columns
    'wshBaseHours.Range("A:P").CurrentRegion.EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    
End Sub

Sub TEC_FilterAndSort()
    'You need the two Non Null Values to Filter
    If wshBaseHours.Range("S3").value = "" Or wshBaseHours.Range("T3").value = "" Then
        Exit Sub
    End If
    
    ImportTEC '2023-12-15 @ 17:02
    
    With wshBaseHours
        Dim LastRow As Long, LastResultRow As Long, ResultRow As Long
        LastRow = .Range("A999999").End(xlUp).Row 'Last BaseHours Row
        If LastRow < 2 Then Exit Sub 'Nothing to filter
        Application.ScreenUpdating = False
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        .Range("A2:Q" & LastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("S2:U3"), _
            CopyToRange:=.Range("W2:AJ2"), _
            Unique:=True
        LastResultRow = .Range("W999999").End(xlUp).Row
        If LastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If LastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wshBaseHours.Range("Y3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add Key:=wshBaseHours.Range("W3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SetRange wshBaseHours.Range("W3:AJ" & LastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
NoSort:
    End With
    Application.ScreenUpdating = True
End Sub

'************************************************************** EffaceFormulaire
Sub EffaceFormulaire()

    'Empty the dynamic fields after reseting the form
    With frmSaisieHeures
        .txtClient.value = ""
        wshAdmin.Range("Client_ID_Admin").value = 0
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .cmbProfessionnel.Enabled = True
        .txtDate.Enabled = True
    End With
    
    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    With frmSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdDelete.Enabled = False
        .cmdUpdate.Enabled = False
    End With
        
    frmSaisieHeures.txtClient.SetFocus
    
End Sub

'************************************************************* AjouteLigneDetail
Sub AjouteLigneDetail()

    If IsDataValid() = False Then Exit Sub
    
    AddOrUpdateTECRecordToDB (0) 'Write to external XLSX file - 2023-12-15 @ 17:09

    'Empty the fields after saving
    frmSaisieHeures.txtClient.value = ""
    frmSaisieHeures.txtActivite.value = ""
    frmSaisieHeures.txtHeures.value = ""
    frmSaisieHeures.txtCommNote.value = ""
        
    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    'Change buttons status
    With frmSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdUpdate.Enabled = False
    End With
    
    frmSaisieHeures.txtClient.SetFocus
    
End Sub

Sub AddOrUpdateTECRecordToDB(r As Long) '2023-12-15 @ 13:33
    Dim FullFileName As String
    Dim SheetName As String
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim MaxID As Long
    Dim LastRow As Long
    Dim nextID As Long
    
    Application.ScreenUpdating = False
    
    Debug.Print "Dans AddOrUpdateTECRecordToDB, r vaut " & r
    
    FullFileName = wshAdmin.Range("SharedFolder").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    SheetName = "TEC"
    
    'Initialize connection, connection string & open the connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Set rs = CreateObject("ADODB.Recordset")

    'If r is 0, add a new record; otherwise, update an existing record
    If r = 0 Then
    'SQL select command to find the next available ID
        strSQL = "SELECT MAX(TEC_ID) AS MaxID FROM [" & SheetName & "$]"
    
        'Open recordset to find out the MaxID
        rs.Open strSQL, conn
        
        'Get the last used row
        If IsNull(rs.Fields("MaxID").value) Then
            ' Handle empty table (assign a default value, e.g., 1)
            LastRow = 1
        Else
            LastRow = rs.Fields("MaxID").value
        End If
        
        'Calculate the new ID
        nextID = LastRow + 1
    
        'Close the previous recordset, no longer needed and open an empty recordset
        rs.Close
        rs.Open "SELECT * FROM [" & SheetName & "$] WHERE 1=0", conn, 2, 3
        rs.AddNew
        
        'Add fields to the recordset before updating it
        rs.Fields("TEC_ID").value = nextID
        rs.Fields("Prof_ID").value = wshAdmin.Range("Prof_ID")
        rs.Fields("Prof").value = frmSaisieHeures.cmbProfessionnel.value
        rs.Fields("Date").value = CDate(frmSaisieHeures.txtDate.value)
        rs.Fields("Client_ID").value = wshAdmin.Range("Client_ID_Admin")
        rs.Fields("ClientNom").value = frmSaisieHeures.txtClient.value
        rs.Fields("Description").value = frmSaisieHeures.txtActivite.value
        rs.Fields("Heures").value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
        rs.Fields("CommentaireNote").value = frmSaisieHeures.txtCommNote.value
        rs.Fields("EstFacturable").value = frmSaisieHeures.chbFacturable.value
        rs.Fields("DateSaisie").value = Now
        rs.Fields("EstFacturee").value = False
        rs.Fields("DateFacturee").value = ""
        rs.Fields("EstDetruit").value = False
        rs.Fields("VersionApp").value = gAppVersion
        rs.Fields("NoFacture").value = ""
        rs.Fields("BD_Row").value = 0 'TO-DO - Remove that column, no longer required - 2023-12-16
    Else
        'If r is not 0, update an existing record (Only fields that can be different)
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & SheetName & "$] WHERE TEC_ID=" & r, conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
'            rs.Fields("TEC_ID").value = nextID
'            rs.Fields("Prof_ID").value = wshAdmin.Range("Prof_ID")
'            rs.Fields("Prof").value = frmSaisieHeures.cmbProfessionnel.value
'            rs.Fields("Date").value = CDate(frmSaisieHeures.txtDate.value)
            rs.Fields("Client_ID").value = wshAdmin.Range("Client_ID_Admin")
            rs.Fields("ClientNom").value = frmSaisieHeures.txtClient.value
            rs.Fields("Description").value = frmSaisieHeures.txtActivite.value
            rs.Fields("Heures").value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
            rs.Fields("CommentaireNote").value = frmSaisieHeures.txtCommNote.value
            rs.Fields("EstFacturable").value = frmSaisieHeures.chbFacturable.value
            rs.Fields("DateSaisie").value = Now
'            rs.Fields("EstFacturee").value = False
'            rs.Fields("DateFacturee").value = ""
'            rs.Fields("EstDetruit").value = False
            rs.Fields("VersionApp").value = gAppVersion
'            rs.Fields("NoFacture").value = ""
'            rs.Fields("BD_Row").value = 0 'TO-DO - Remove that column, no longer required - 2023-12-16
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut être trouvé!", vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    End If

    'Update the recordset (create the record)
    rs.Update
    rs.Close
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

'************************************************************ ModifieLigneDetail
Sub ModifieLigneDetail()

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox Prompt:="Vous devez choisir un enregistrement à modifier !", _
               Title:="", _
               Buttons:=vbCritical
        Exit Sub
    End If
    
    If IsDataValid() = False Then Exit Sub

    Dim sh As Worksheet
    Set sh = wshBaseHours

    Dim selectedRow As Long
    Debug.Print "Le ID du record à réécrire, selon Admin est '" & wshAdmin.Range("TEC_Current_ID").value & "'"
    selectedRow = wshAdmin.Range("TEC_Current_ID").value
    Debug.Print "Le ID du record à réécrire est '" & selectedRow & "'"
    
    'WriteToWorksheet (selectedRow)
    AddOrUpdateTECRecordToDB (selectedRow) 'Write to external XLSX file - 2023-12-15 @ 13:33
 
    frmSaisieHeures.txtClient.value = ""
    frmSaisieHeures.txtActivite.value = ""
    frmSaisieHeures.txtHeures.value = ""
    frmSaisieHeures.txtCommNote.value = ""

    frmSaisieHeures.cmbProfessionnel.Enabled = True
    frmSaisieHeures.txtDate.Enabled = True
    rmv_state = rmv_modeCreation

    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    frmSaisieHeures.txtClient.SetFocus

End Sub

'************************************************************* EffaceLigneDetail
Sub EffaceLigneDetail()

    If frmSaisieHeures.txtID.value = "" Then
        MsgBox _
        Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
        Title:="", _
        Buttons:=vbCritical
        Exit Sub
    End If
    
    Dim answerYesNo As Integer
    answerYesNo = MsgBox("Êtes-vous certain de vouloir DÉTRUIRE cet enregistrement ? ", _
                         vbYesNo + vbQuestion, "Confirmation de DESTRUCTION")
    If answerYesNo = vbNo Then
        MsgBox _
        Prompt:="Cet enregistrement ne sera PAS détruit ! ", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("HeuresBase")
    
    Dim selectedRow As Long
    selectedRow = Application.WorksheetFunction.Match(CLng(frmSaisieHeures.txtID.value), _
                                                      sh.Range("A:A"), 0)
    
    'Assign 'VRAI' to colomn 14, since it is deleted
    sh.Range("K" & selectedRow).value = Now
    sh.Range("N" & selectedRow).value = True
    sh.Range("O" & selectedRow).value = gAppVersion
    
    'Empty the dynamic fields after deleting
    With frmSaisieHeures
        .txtClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
    End With
    
    MsgBox _
        Prompt:="L'enregistrement a été DÉTRUIT !", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        
    frmSaisieHeures.cmbProfessionnel.Enabled = True
    frmSaisieHeures.txtDate.Enabled = True
    rmv_state = rmv_modeCreation
    
    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    frmSaisieHeures.txtClient.SetFocus

End Sub

Function IsDataValid() As Boolean

    IsDataValid = False
    
    'Validations first (one field at a time)
    If frmSaisieHeures.cmbProfessionnel.value = "" Then
        MsgBox Prompt:="Le professionnel est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.cmbProfessionnel.SetFocus
        Exit Function
    End If

    If frmSaisieHeures.txtDate.value = "" Or IsDate(frmSaisieHeures.txtDate.value) = False Then
        MsgBox Prompt:="La date est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.txtDate.SetFocus
        Exit Function
    End If

    If frmSaisieHeures.txtClient.value = "" Then
        MsgBox Prompt:="Le client est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.txtClient.SetFocus
        Exit Function
    End If
    
    If frmSaisieHeures.txtHeures.value = "" Or IsNumeric(frmSaisieHeures.txtHeures.value) = False Then
        MsgBox Prompt:="Le nombre d'heures est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.txtHeures.SetFocus
        Exit Function
    End If

    IsDataValid = True

End Function

Sub WriteToWorksheet(r As Long)

    'Load the cmb & txt into the 'HeuresBase' worksheet
    With wshBaseHours
        .Range("A" & r).value = LastRow
        .Range("B" & r).value = wshAdmin.Range("Prof_ID")
        .Range("C" & r).value = frmSaisieHeures.cmbProfessionnel.value
        .Range("D" & r).value = CDate(frmSaisieHeures.txtDate.value)
        .Range("E" & r).value = wshAdmin.Range("Client_ID_Admin")
        .Range("F" & r).value = frmSaisieHeures.txtClient.value
        .Range("G" & r).value = frmSaisieHeures.txtActivite.value
        .Range("H" & r).value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
        .Range("I" & r).value = frmSaisieHeures.txtCommNote.value
        .Range("J" & r).value = frmSaisieHeures.chbFacturable.value
        .Range("K" & r).value = Now
        .Range("L" & r).value = False
        .Range("M" & r).value = ""
        .Range("N" & r).value = False
        .Range("O" & r).value = gAppVersion
        .Range("P" & r).value = ""
    End With

    MsgBox "Record has been added to wshBaseHours - " & LastRow
    
End Sub

'********************* Reload listBox from HeuresFiltered and reset the buttons
Sub RefreshListBoxAndAddHours()

    If wshAdmin.Range("Prof_ID").value = "" Or wshAdmin.Range("TECDate").value = "" Then
        GoTo EndOfProcedure
    End If
    
    frmSaisieHeures.txtTotalHeures.value = ""
    
    Dim shFiltered As Worksheet
    Set shFiltered = ThisWorkbook.Sheets("HeuresBase")
    'shFiltered.Activate
    
    'Last Row used in column A
    Dim LastRow As Long
    LastRow = wshBaseHours.Range("W99999").End(xlUp).Row - 1
    If LastRow = 0 Then Exit Sub
        
    With frmSaisieHeures.lstData
        .ColumnHeads = True
        .ColumnCount = 9
        .ColumnWidths = "28; 26; 51; 130; 180; 35; 80; 32; 83"
        
        If LastRow = 1 Then
            .RowSource = "HeuresBase!W3:AE3"
        Else
            .RowSource = "HeuresBase!W3:AE" & LastRow + 1
        End If
    End With

    'Add hours to totalHeures
    Dim nbrRows, i As Integer
    nbrRows = frmSaisieHeures.lstData.ListCount
    Dim totalHeures As Double
    
    If nbrRows > 0 Then
        For i = 0 To nbrRows - 1
            totalHeures = totalHeures + CCur(frmSaisieHeures.lstData.List(i, 5))
        Next
        frmSaisieHeures.txtTotalHeures.value = Format(totalHeures, "#0.00")
    End If

EndOfProcedure:
    frmSaisieHeures.cmdClear.Enabled = False
    frmSaisieHeures.cmdAdd.Enabled = False
    frmSaisieHeures.cmdUpdate.Enabled = False
    frmSaisieHeures.cmdDelete.Enabled = False

    'frmSaisieHeures.txtClient.SetFocus
    
End Sub

'Sub AddNewRecordToClosedFile() '2023-12-15 @ 11:40
'    Dim conn As Object
'    Dim rs As Object
'    Dim strSQL As String
'    Dim FullFileName As String
'    Dim SheetName As String
'    Dim LastRow As Long
'    Dim newID As Long
'
'    Application.ScreenUpdating = False
'
'    FullFileName = wshAdmin.Range("SharedFolder").value & Application.PathSeparator & _
'                   "GCF_DB_Sortie.xlsx"
'    SheetName = "TEC"
'
'    ' Set up connection
'    Set conn = CreateObject("ADODB.Connection")
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
'
'    'Set up recordset
'    Set rs = CreateObject("ADODB.Recordset")
'
'    'SQL query to get the last used row in column A (adjust column as needed)
'    strSQL = "SELECT MAX(TEC_ID) AS LastRow FROM [" & SheetName & "$]"
'
'    'Open recordset
'    rs.Open strSQL, conn
'
'    'Get the last used row
'    If IsNull(rs.Fields("LastRow").value) Then
'        ' Handle empty table (assign a default value, e.g., 1)
'        LastRow = 1
'    Else
'        LastRow = rs.Fields("LastRow").value
'    End If
'
'    'Calculate the new ID
'    newID = LastRow + 1
'
'    'Close the previous recordset, no longer needed and open an empty recordset
'    rs.Close
'    rs.Open "SELECT * FROM [" & SheetName & "$] WHERE 1=0", conn, 2, 3
'    rs.AddNew
'
'    rs.Fields("TEC_ID").value = newID
'    Debug.Print wshAdmin.Range("Prof_ID") & " - " & wshAdmin.Range("B4").value
'    rs.Fields("Prof_ID").value = wshAdmin.Range("Prof_ID")
'    rs.Fields("Prof").value = frmSaisieHeures.cmbProfessionnel.value
'    rs.Fields("Date").value = CDate(frmSaisieHeures.txtDate.value)
'    rs.Fields("Client_ID").value = wshAdmin.Range("Client_ID_Admin")
'    rs.Fields("ClientNom").value = frmSaisieHeures.txtClient.value
'    rs.Fields("Description").value = frmSaisieHeures.txtActivite.value
'    rs.Fields("Heures").value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
'    rs.Fields("CommentaireNote").value = frmSaisieHeures.txtCommNote.value
'    rs.Fields("EstFacturable").value = frmSaisieHeures.chbFacturable.value
'    rs.Fields("DateSaisie").value = Now
'    rs.Fields("EstFacturee").value = False
'    rs.Fields("DateFacturee").value = ""
'    rs.Fields("EstDetruit").value = False
'    rs.Fields("VersionApp").value = gAppVersion
'    rs.Fields("NoFacture").value = ""
'
'    'Update the recordset
'    rs.Update
'
'    'Display a message indicating the new record has been added
'    MsgBox "New record with ID " & newID & " added to the closed file."
'
'    ' Close connections
'    rs.Close
'    Set rs = Nothing
'    conn.Close
'    Set conn = Nothing
'
'    Application.ScreenUpdating = True
'
'End Sub
'
'Sub AddTECRecordToDB() '2023-12-15 @ 11:03
'    Dim FullFileName As String
'    Dim SheetName As String
'    Dim conn As Object
'    Dim rs As Object
'    Dim strConn As String
'    Dim strSQL As String
'    Dim MaxID As Long
'    Dim nextID As Long
'
'    Application.ScreenUpdating = False
'
'    FullFileName = wshAdmin.Range("SharedFolder").value & Application.PathSeparator & _
'                   "GCF_DB_Sortie.xlsx"
'    SheetName = "TEC"
'
'    'Initialize connection, connection string & open the connection
'    Set conn = CreateObject("ADODB.Connection")
'    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
'    conn.Open strConn
'
'    'Initialize recordset
'    Set rs = CreateObject("ADODB.Recordset")
'
'    'SQL select command to find the next available ID
'    strSQL = "SELECT MAX(TEC_ID) AS MaxID FROM [" & SheetName & "$]"
'
'    'Open the recordset with the select command
'    rs.Open strSQL, conn, 2, 3
'
'    'Check if the recordset is not empty
'    If rs.Fields("MaxID").value <> vbNull Then
'        nextID = rs.Fields("MaxID").value + 1
'    Else
'        nextID = 1
'    End If
'
'    'Close the previous recordset, no longer needed and open an empty recordset
'    rs.Close
'    rs.Open "SELECT * FROM [" & SheetName & "$] WHERE 1=0", conn, 2, 3
'    rs.AddNew
'
'    'Add fields to the recordset before updating it
'    rs.Fields("TEC_ID").value = nextID
'    Debug.Print wshAdmin.Range("Prof_ID") & " - " & wshAdmin.Range("B4").value
'    rs.Fields("Prof_ID").value = wshAdmin.Range("Prof_ID")
'    rs.Fields("Prof").value = frmSaisieHeures.cmbProfessionnel.value
'    rs.Fields("Date").value = CDate(frmSaisieHeures.txtDate.value)
'    rs.Fields("Client_ID").value = wshAdmin.Range("Client_ID_Admin")
'    rs.Fields("ClientNom").value = frmSaisieHeures.txtClient.value
'    rs.Fields("Description").value = frmSaisieHeures.txtActivite.value
'    rs.Fields("Heures").value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
'    rs.Fields("CommentaireNote").value = frmSaisieHeures.txtCommNote.value
'    rs.Fields("EstFacturable").value = frmSaisieHeures.chbFacturable.value
'    rs.Fields("DateSaisie").value = Now
'    rs.Fields("EstFacturee").value = False
'    rs.Fields("DateFacturee").value = ""
'    rs.Fields("EstDetruit").value = False
'    rs.Fields("VersionApp").value = gAppVersion
'    rs.Fields("NoFacture").value = ""
'
'    'Update the recordset (create the record)
'    rs.Update
'    rs.Close
'
'    'Close recordset and connection
'    On Error Resume Next
'    rs.Close
'    On Error GoTo 0
'    conn.Close
'
'    Application.ScreenUpdating = True
'
'End Sub

'Sub TestADOQuery()
'    Dim conn As Object
'    Dim rs As Object
'    Dim strSQL As String
'
'    Dim FullFileName As String
'    FullFileName = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_DB_Sortie.xlsx"
'    Dim SheetName As String
'    SheetName = "TEC"
'
'    'Create connection
'    Set conn = CreateObject("ADODB.Connection")
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
'
'    'Create recordset
'    Set rs = CreateObject("ADODB.Recordset")
'
'    ' SQL Query with Tec_ID included
'    strSQL = "SELECT MAX(Tec_ID) AS MaxID FROM [" & SheetName & "$]"
'
'    ' Execute the query
'    rs.Open strSQL, conn
'
'    'Check if the recordset is not empty
'    If rs.Fields("MaxID").value <> vbNull Then
'        MaxID = rs.Fields("MaxID").value + 1
'    Else
'        MaxID = 1
'    End If
'
'    ' Close connections
'    rs.Close
'    Set rs = Nothing
'    conn.Close
'    Set conn = Nothing
'
'End Sub

'Sub TestADOQuery2()
'    Dim conn As Object
'    Dim rs As Object
'    Dim strSQL As String
'    Dim MaxID As Long
'
'    Dim FullFileName As String
'    FullFileName = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_DB_Sortie.xlsx"
'    Dim SheetName As String
'    SheetName = "TEC"
'
'    ' Create connection
'    Set conn = CreateObject("ADODB.Connection")
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullFileName & ";Extended Properties=""Excel 12.0;HDR=YES"";"
'
'    ' Create recordset
'    Set rs = CreateObject("ADODB.Recordset")
'
'    ' SQL Query with Tec_ID included
'    strSQL = "SELECT MAX(Tec_ID) AS MaxID FROM [" & SheetName & "$]"
'
'    ' Execute the query
'    rs.Open strSQL, conn
'
'    'Check if the recordset is not empty
'    If rs.Fields("MaxID").value <> vbNull Then
'        MaxID = rs.Fields("MaxID").value + 1
'    Else
'        MaxID = 1
'    End If
'
'    ' Close connections
'    rs.Close
'    Set rs = Nothing
'    conn.Close
'    Set conn = Nothing
'
'End Sub

