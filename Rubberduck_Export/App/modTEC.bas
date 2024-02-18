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

Global Const gAppVersion As String = "v2.2" '2024-02-14 @ 09:31

Sub Client_List_Import_All() 'Using ADODB - 2024-02-14 @ 07:22
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshClientDB.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim fullFileName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & _
                   Application.PathSeparator & "GCF_BD_Entrée.xlsx" '2024-02-14 @ 07:04
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = fullFileName
    sourceTab = "Clients"
    
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
    recSet.Source = "SELECT * FROM [" & sourceTab & "$]"
    recSet.Open
    
    'Copy to wshClientDB workbook
    wshClientDB.Range("A2").CopyFromRecordset recSet
    wshClientDB.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
'    MsgBox _
'        Prompt:="J'ai importé un total de " & _
'            Format(wshClientDB.Range("A1").CurrentRegion.Rows.count - 1, _
'            "## ##0") & " clients", _
'        Title:="Vérification du nombre de clients", _
'        Buttons:=vbInformation
        
End Sub

Sub TEC_Import_All() '2024-02-14 @ 06:19
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshBaseHours.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim fileName As String, sourceWorkbook As String, sourceTab As String
    fileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                "GCF_BD_Sortie.xlsx" '2024-02-14 @ 06:22
    sourceWorkbook = fileName
    sourceTab = "TEC"
    
    'Set up source and destination ranges
    Dim sourceRange As Range, destinationRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).UsedRange
    Set destinationRange = wshBaseHours.Range("A2")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshBaseHours.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Arrange formats on all rows
    Dim LastRow As Long
    LastRow = wshBaseHours.Range("A999999").End(xlUp).row
    
    With wshBaseHours
        .Range("A3" & ":P" & LastRow).HorizontalAlignment = xlCenter
        With .Range("F3:F" & LastRow & ",G3:G" & LastRow & ",I3:I" & LastRow & ",O3:O" & LastRow)
            .HorizontalAlignment = xlLeft
        End With
        .Range("H3:H" & LastRow).NumberFormat = "#0.00"
        .Range("K3:K" & LastRow).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Sub TEC_Advanced_Filter_And_Sort() '2024-02-14 @ 06:41
    'Two Non Null Values are mandatory to Advanced Filter
    If wshBaseHours.Range("R3").value = "" Or wshBaseHours.Range("S3").value = "" Then
        Exit Sub
    End If
    
    Call TEC_Import_All '2024-02-14 @ 06:20
    
    With wshBaseHours
        Dim LastRow As Long, LastResultRow As Long, ResultRow As Long
        LastRow = .Range("A999999").End(xlUp).row 'Last BaseHours Row
        If LastRow < 3 Then Exit Sub 'Nothing to filter
        Application.ScreenUpdating = False
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        'Advanced Filter applied to BaseHours
        .Range("A2:P" & LastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("R2:W3"), _
            CopyToRange:=.Range("Y2:AL2"), _
            Unique:=True
        'Analyze Advance Filter Results
        LastResultRow = .Range("Y999999").End(xlUp).row
        If LastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=wshBaseHours.Range("AA3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add key:=wshBaseHours.Range("Y3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Tec_ID
            .SetRange wshBaseHours.Range("Y3:AL" & LastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
No_Sort_Required:
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
    
    Call TEC_Advanced_Filter_And_Sort
    Call Refresh_ListBox_And_Add_Hours
    
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
    
    Call Add_Or_Update_TEC_Record_To_DB(0) 'Write to external XLSX file - 2023-12-23 @ 07:03
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

    Call TEC_Advanced_Filter_And_Sort
    Call Refresh_ListBox_And_Add_Hours
    
    'Reset command buttons
    With frmSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdUpdate.Enabled = False
    End With
    
    'Back to client
    frmSaisieHeures.txtClient.SetFocus
    
End Sub

Sub ModifieLigneDetail() '2023-12-23 @ 07:04

    If IsDataValid() = False Then Exit Sub

    Add_Or_Update_TEC_Record_To_DB (wshAdmin.Range("TEC_Current_ID").value) 'Write to external XLSX file - 2023-12-16 @ 14:10
 
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

    Call TEC_Advanced_Filter_And_Sort
    Call Refresh_ListBox_And_Add_Hours
    
    rmv_state = rmv_modeCreation
    
    frmSaisieHeures.txtClient.SetFocus

End Sub

Sub EffaceLigneDetail() '2023-12-23 @ 07:05

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox _
        Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
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
    Set sh = wshBaseHours
    
    Dim selectedRow As Long
    'With a negative ID value, it means to soft delete this record
    selectedRow = -wshAdmin.Range("TEC_Current_ID").value
    Add_Or_Update_TEC_Record_To_DB (selectedRow) 'Write to external XLSX file - 2023-12-23 @ 07:07
    
    'Empty the dynamic fields after deleting
    With frmSaisieHeures
        .txtClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .chbFacturable = True
    End With
    
    MsgBox _
        Prompt:="L'enregistrement a été DÉTRUIT !", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        
    frmSaisieHeures.cmbProfessionnel.Enabled = True
    frmSaisieHeures.txtDate.Enabled = True
    rmv_state = rmv_modeCreation
    
    Call TEC_Advanced_Filter_And_Sort
    Call Refresh_ListBox_And_Add_Hours
    
    frmSaisieHeures.txtClient.SetFocus

End Sub

Sub Add_Or_Update_TEC_Record_To_DB(r As Long) 'Write -OR- Update a record to external .xlsx file
    
    Application.ScreenUpdating = False
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    sheetName = "TEC"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    If r < 0 Then 'Soft delete a record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & sheetName & "$] WHERE TEC_ID=" & Abs(r), conn, 2, 3
        If Not rs.EOF Then
            'Update the "IsDeleted" field to mark the record as deleted
            rs.Fields("DateSaisie").value = Now
            rs.Fields("EstDetruit").value = True
            rs.Fields("VersionApp").value = gAppVersion
            rs.Update
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut être trouvé!", _
                vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    Else
        'If r is 0, add a new record; otherwise, update an existing record
        If r = 0 Then 'Add a record
        'SQL select command to find the next available ID
            Dim strSQL As String, MaxID As Long
            strSQL = "SELECT MAX(TEC_ID) AS MaxID FROM [" & sheetName & "$]"
        
            'Open recordset to find out the MaxID
            rs.Open strSQL, conn
            
            'Get the last used row
            Dim LastRow As Long
            If IsNull(rs.Fields("MaxID").value) Then
                ' Handle empty table (assign a default value, e.g., 1)
                LastRow = 1
            Else
                LastRow = rs.Fields("MaxID").value
            End If
            
            'Calculate the new ID
            Dim nextID As Long
            nextID = LastRow + 1
        
            'Close the previous recordset, no longer needed and open an empty recordset
            rs.Close
            rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
            
            'Add fields to the recordset before updating it
            rs.AddNew
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
            rs.Open "SELECT * FROM [" & sheetName & "$] WHERE TEC_ID=" & r, conn, 2, 3
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
                MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut être trouvé!", vbExclamation
                rs.Close
                conn.Close
                Exit Sub
            End If
        End If
    End If
    'Update the recordset (create the record)
    rs.Update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

Sub Refresh_ListBox_And_Add_Hours() 'Load the listBox with the appropriate records

    If wshAdmin.Range("TEC_Prof_ID").value = "" Or wshAdmin.Range("TEC_Date").value = "" Then
        GoTo EndOfProcedure
    End If
    
    frmSaisieHeures.txtTotalHeures.value = ""
    
    'Last Row used in first column of result
    Dim LastRow As Long
    LastRow = wshBaseHours.Range("Y99999").End(xlUp).row - 1
    If LastRow = 0 Then Exit Sub
        
    With frmSaisieHeures.lstData
        .ColumnHeads = True
        .ColumnCount = 9
        .ColumnWidths = "28; 26; 51; 130; 180; 35; 80; 32; 83"
        
        If LastRow = 1 Then
            .RowSource = "HeuresBase!Y3:AG3"
        Else
            .RowSource = "HeuresBase!Y3:AG" & LastRow + 1
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

