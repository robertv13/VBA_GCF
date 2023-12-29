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

Global Const gAppVersion As String = "v1.1.9.C"

Sub ImportClientList() '2023-11-23 @ 06:51
    
    'Clear all cells, but the headers, in the target worksheet
    wshClientDB.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_BD_Entr�e.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx" '2023-12-23 06:53
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
    wshClientDB.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
'    MsgBox _
'        Prompt:="J'ai import� un total de " & _
'            Format(wshClientDB.Range("A1").CurrentRegion.Rows.count - 1, _
'            "## ##0") & " clients", _
'        Title:="V�rification du nombre de clients", _
'        Buttons:=vbInformation
        
End Sub

Sub TEC_Import()
    
    Dim startTime As String
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the target worksheet
    wshBaseHours.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2023-12-23 @ 06:54

    'Set up source and destination ranges
    Dim sourceRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets("TEC").UsedRange

    Dim destinationRange As Range
    Set destinationRange = wshBaseHours.Range("A2")

    'Copy data, using Range to Range
    sourceRange.Copy destinationRange
    wshBaseHours.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim lastRow As Long
    lastRow = wshBaseHours.Range("A999999").End(xlUp).row
    
    With wshBaseHours
        With .Range("A3" & ":P" & lastRow)
            .HorizontalAlignment = xlCenter
        End With
        With .Range("F3:F" & lastRow & ",G3:G" & lastRow & ",I3:I" & lastRow & ",O3:O" & lastRow)
            .HorizontalAlignment = xlLeft
        End With
        .Range("H3:H" & lastRow).NumberFormat = "#0.00"
        .Range("K3:K" & lastRow).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Sub TEC_FilterAndSort()
    'You need the two Non Null Values to Filter
    If wshBaseHours.Range("R3").value = "" Or wshBaseHours.Range("S3").value = "" Then
        Exit Sub
    End If
    
    TEC_Import '2023-12-23 @ 06:58
    
    With wshBaseHours
        Dim lastRow As Long, lastResultRow As Long, ResultRow As Long
        lastRow = .Range("A999999").End(xlUp).row 'Last BaseHours Row
        If lastRow < 3 Then Exit Sub 'Nothing to filter
        Application.ScreenUpdating = False
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        'Advanced Filter applied to BaseHours
        .Range("A2:Q" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("R2:W3"), _
            CopyToRange:=.Range("Y2:AL2"), _
            Unique:=True
        lastResultRow = .Range("Y999999").End(xlUp).row
        If lastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wshBaseHours.Range("Y3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SortFields.Add Key:=wshBaseHours.Range("Z3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SetRange wshBaseHours.Range("Y3:AL" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
NoSort:
    End With
    Application.ScreenUpdating = True
End Sub

Sub EffaceFormulaire() 'Clear all fields on the userForm

    'Empty the dynamic fields after reseting the form
    With frmSaisieHeures
        .txtClient.value = ""
        wshAdmin.Range("TEC_Client_ID").value = 0
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

Sub AjouteLigneDetail() 'Add an entry to DB

    If IsDataValid() = False Then Exit Sub
    
    AddOrUpdateTECRecordToDB (0) 'Write to external XLSX file - 2023-12-23 @ 07:03
    'Clear the fields after saving
    With frmSaisieHeures
'        .cmbProfessionnel.Enabled = True
'        .txtDate.Enabled = True
        .txtClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .chbFacturable = True
    End With

    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    'Reset command buttons
    With frmSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdUpdate.Enabled = False
    End With
    
    frmSaisieHeures.txtClient.SetFocus
    
End Sub

Sub ModifieLigneDetail() '2023-12-23 @ 07:04

    If IsDataValid() = False Then Exit Sub

    AddOrUpdateTECRecordToDB (wshAdmin.Range("TEC_Current_ID").value) 'Write to external XLSX file - 2023-12-16 @ 14:10
 
    'Initialize dynamic variables
    With frmSaisieHeures
        .cmbProfessionnel.Enabled = True
        .txtDate.Enabled = True
        .txtClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .chbFacturable = True
    End With

    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    rmv_state = rmv_modeCreation
    
    frmSaisieHeures.txtClient.SetFocus

End Sub

Sub EffaceLigneDetail() '2023-12-23 @ 07:05

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox _
        Prompt:="Vous devez choisir un enregistrement � D�TRUIRE !", _
        Buttons:=vbCritical
        Exit Sub
    End If
    
    Dim answerYesNo As Integer
    answerYesNo = MsgBox("�tes-vous certain de vouloir D�TRUIRE cet enregistrement ? ", _
                         vbYesNo + vbQuestion, "Confirmation de DESTRUCTION")
    If answerYesNo = vbNo Then
        MsgBox _
        Prompt:="Cet enregistrement ne sera PAS d�truit ! ", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = wshBaseHours
    
    Dim selectedRow As Long
    'With a negative ID value, it means to soft delete this record
    selectedRow = -wshAdmin.Range("TEC_Current_ID").value
    AddOrUpdateTECRecordToDB (selectedRow) 'Write to external XLSX file - 2023-12-23 @ 07:07
    
    'Empty the dynamic fields after deleting
    With frmSaisieHeures
        .txtClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .chbFacturable = True
    End With
    
    MsgBox _
        Prompt:="L'enregistrement a �t� D�TRUIT !", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        
    frmSaisieHeures.cmbProfessionnel.Enabled = True
    frmSaisieHeures.txtDate.Enabled = True
    rmv_state = rmv_modeCreation
    
    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    frmSaisieHeures.txtClient.SetFocus

End Sub

Sub AddOrUpdateTECRecordToDB(r As Long) 'Write/Update a record to external .xlsx file
    Dim FullFileName As String
    Dim SheetName As String
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim maxID As Long
    Dim lastRow As Long
    Dim nextID As Long
    
    Application.ScreenUpdating = False
    
    FullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    SheetName = "TEC"
    
    'Initialize connection, connection string & open the connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Set rs = CreateObject("ADODB.Recordset")

    If r < 0 Then 'Soft delete
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & SheetName & "$] WHERE TEC_ID=" & Abs(r), conn, 2, 3
        If Not rs.EOF Then
            'Update the "IsDeleted" field to mark the record as deleted
            rs.Fields("DateSaisie").value = Now
            rs.Fields("EstDetruit").value = True
            rs.Fields("VersionApp").value = gAppVersion
            rs.Update
        Else
            ' Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut �tre trouv�!", vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    Else
        'If r is 0, add a new record; otherwise, update an existing record
        If r = 0 Then 'Add a record
        'SQL select command to find the next available ID
            strSQL = "SELECT MAX(TEC_ID) AS MaxID FROM [" & SheetName & "$]"
        
            'Open recordset to find out the MaxID
            rs.Open strSQL, conn
            
            'Get the last used row
            If IsNull(rs.Fields("MaxID").value) Then
                ' Handle empty table (assign a default value, e.g., 1)
                lastRow = 1
            Else
                lastRow = rs.Fields("MaxID").value
            End If
            
            'Calculate the new ID
            nextID = lastRow + 1
        
            'Close the previous recordset, no longer needed and open an empty recordset
            rs.Close
            rs.Open "SELECT * FROM [" & SheetName & "$] WHERE 1=0", conn, 2, 3
            rs.AddNew
            
            'Add fields to the recordset before updating it
            rs.Fields("TEC_ID").value = nextID
            rs.Fields("Prof_ID").value = wshAdmin.Range("TEC_Prof_ID")
            rs.Fields("Prof").value = frmSaisieHeures.cmbProfessionnel.value
            rs.Fields("Date").value = CDate(frmSaisieHeures.txtDate.value)
            rs.Fields("Client_ID").value = wshAdmin.Range("TEC_Client_ID")
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
        Else 'Update an existing record
            'Open the recordset for the specified ID
            rs.Open "SELECT * FROM [" & SheetName & "$] WHERE TEC_ID=" & r, conn, 2, 3
            If Not rs.EOF Then
                'Update fields for the existing record
                rs.Fields("Client_ID").value = wshAdmin.Range("TEC_Client_ID")
                rs.Fields("ClientNom").value = frmSaisieHeures.txtClient.value
                rs.Fields("Description").value = frmSaisieHeures.txtActivite.value
                rs.Fields("Heures").value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
                rs.Fields("CommentaireNote").value = frmSaisieHeures.txtCommNote.value
                rs.Fields("EstFacturable").value = frmSaisieHeures.chbFacturable.value
                rs.Fields("DateSaisie").value = Now
                rs.Fields("VersionApp").value = gAppVersion
            Else
                'Handle the case where the specified ID is not found
                MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut �tre trouv�!", vbExclamation
                rs.Close
                conn.Close
                Exit Sub
            End If
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

Sub RefreshListBoxAndAddHours() 'Load the listBox with the appropriate records

    If wshAdmin.Range("TEC_Prof_ID").value = "" Or wshAdmin.Range("TEC_Date").value = "" Then
        GoTo EndOfProcedure
    End If
    
    frmSaisieHeures.txtTotalHeures.value = ""
    
    'Last Row used in first column of result
    Dim lastRow As Long
    lastRow = wshBaseHours.Range("Y99999").End(xlUp).row - 1
    If lastRow = 0 Then Exit Sub
        
    With frmSaisieHeures.lstData
        .ColumnHeads = True
        .ColumnCount = 9
        .ColumnWidths = "28; 26; 51; 130; 180; 35; 80; 32; 83"
        
        If lastRow = 1 Then
            .RowSource = "HeuresBase!Y3:AG3"
        Else
            .RowSource = "HeuresBase!Y3:AG" & lastRow + 1
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

