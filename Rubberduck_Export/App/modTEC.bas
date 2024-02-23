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

Global Const gAppVersion As String = "v2.4" '2024-02-23 @ 09:11

Public TabOrderFlag As Boolean 'To be able to specify the TAB order of a worksheet

Sub Client_List_Import_All() 'Using ADODB - 2024-02-14 @ 07:22
    
    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer
    
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
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    Debug.Print vbNewLine & String(60, "*") & vbNewLine & _
        "Client_List_Import_All() - After connStr.Open - Secondes = " & Timer - timerStart & _
        vbNewLine & String(60, "*")
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    
    recSet.ActiveConnection = connStr
    recSet.source = "SELECT * FROM [" & sourceTab & "$]"
    recSet.Open
    
    Debug.Print vbNewLine & String(60, "*") & vbNewLine & _
        "Client_List_Import_All() - After recSet.Open - Secondes = " & Timer - timerStart & _
        vbNewLine & String(60, "*")
    
    'Copy to wshClientDB workbook
    wshClientDB.Range("A2").CopyFromRecordset recSet
    
    Debug.Print vbNewLine & String(60, "*") & vbNewLine & _
        "Client_List_Import_All() - After .CopyFromRecordset - Secondes = " & Timer - timerStart & _
        vbNewLine & String(60, "*")

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

    'Free up memory - 2024-02-23
    Set connStr = Nothing
    Set recSet = Nothing

    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        "Client_List_Import_All() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")
        
End Sub

Sub TEC_Import_All() '2024-02-14 @ 06:19
    
    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer
    
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
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange
    Set destinationRange = wshBaseHours.Range("A2")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshBaseHours.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Arrange formats on all rows
    Dim lastRow As Long
    lastRow = wshBaseHours.Range("A99999").End(xlUp).row
    
    With wshBaseHours
        .Range("A3" & ":P" & lastRow).HorizontalAlignment = xlCenter
        With .Range("F3:F" & lastRow & ",G3:G" & lastRow & ",I3:I" & lastRow & ",O3:O" & lastRow)
            .HorizontalAlignment = xlLeft
        End With
        .Range("H3:H" & lastRow).NumberFormat = "#0.00"
        .Range("K3:K" & lastRow).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End With
    
    Application.ScreenUpdating = True
    
    'Free up memory - 2024-02-23
    Set sourceRange = Nothing
    Set destinationRange = Nothing

    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
                "TEC_Import_All() - Secondes = " & Timer - timerStart & _
                vbNewLine & String(45, "*")
    
End Sub

Sub TEC_Advanced_Filter_And_Sort() '2024-02-14 @ 06:41
    
    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer

    'Two Non Null Values are mandatory to Advanced Filter
    If wshBaseHours.Range("R3").value = "" Or wshBaseHours.Range("S3").value = "" Then
        Exit Sub
    End If
    
    Call TEC_Import_All '2024-02-14 @ 06:20
    
    With wshBaseHours
        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
        lastRow = .Range("A999999").End(xlUp).row 'Last BaseHours Row
        If lastRow < 3 Then Exit Sub 'Nothing to filter
        Application.ScreenUpdating = False
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        'Advanced Filter applied to BaseHours
        .Range("A2:P" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("R2:W3"), _
            CopyToRange:=.Range("Y2:AL2"), _
            Unique:=True
        'Analyze Advance Filter Results
        lastResultRow = .Range("Y999999").End(xlUp).row
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wshBaseHours.Range("AB3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add Key:=wshBaseHours.Range("AA3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Prof_ID
            .SortFields.Add Key:=wshBaseHours.Range("Y3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Tec_ID
            .SetRange wshBaseHours.Range("Y3:AL" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
No_Sort_Required:
    End With
    Application.ScreenUpdating = True
    
    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
                "TEC_Import_All() - Secondes = " & Timer - timerStart & _
                vbNewLine & String(45, "*")

End Sub

Sub TEC_Efface_Formulaire() 'Clear all fields on the userForm

    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer

    'Empty the dynamic fields after reseting the form
    With ufSaisieHeures
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
    
    With ufSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdDelete.Enabled = False
        .cmdUpdate.Enabled = False
    End With
        
    ufSaisieHeures.txtClient.SetFocus
    
    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        "TEC_Efface_Formulaire() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")

End Sub

Sub AjouteLigneDetail() 'Add an entry to DB

    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer

    If IsDataValid() = False Then Exit Sub
    
    Call Add_Or_Update_TEC_Record_To_DB(0) 'Write to external XLSX file - 2023-12-23 @ 07:03
    'Clear the fields after saving
    With ufSaisieHeures
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
    With ufSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdUpdate.Enabled = False
    End With
    
    'Back to client
    ufSaisieHeures.txtClient.SetFocus
    
    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        "AjouteLigneDetail() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")

End Sub

Sub ModifieLigneDetail() '2023-12-23 @ 07:04

    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer

    If IsDataValid() = False Then Exit Sub

    Add_Or_Update_TEC_Record_To_DB (wshAdmin.Range("TEC_Current_ID").value) 'Write to external XLSX file - 2023-12-16 @ 14:10
 
    'Initialize dynamic variables
    With ufSaisieHeures
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
    
    ufSaisieHeures.txtClient.SetFocus
    
    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        "ModifieLigneDetail() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")

End Sub

Sub EffaceLigneDetail() '2023-12-23 @ 07:05

    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer

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
    
    Dim sh As Worksheet: Set sh = wshBaseHours
    
    Dim selectedRow As Long
    'With a negative ID value, it means to soft delete this record
    selectedRow = -wshAdmin.Range("TEC_Current_ID").value
    Add_Or_Update_TEC_Record_To_DB (selectedRow) 'Write to external XLSX file - 2023-12-23 @ 07:07
    
    'Empty the dynamic fields after deleting
    With ufSaisieHeures
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
        
    ufSaisieHeures.cmbProfessionnel.Enabled = True
    ufSaisieHeures.txtDate.Enabled = True
    rmv_state = rmv_modeCreation
    
    Call TEC_Advanced_Filter_And_Sort
    Call Refresh_ListBox_And_Add_Hours
    
    ufSaisieHeures.txtClient.SetFocus

    'Free up memory - 2024-02-23
    Set sh = Nothing

    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        "EffaceLigneDetail() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")

End Sub

Sub Add_Or_Update_TEC_Record_To_DB(r As Long) 'Write -OR- Update a record to external .xlsx file
    
    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer

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
            Dim lastRow As Long
            If IsNull(rs.Fields("MaxID").value) Then
                ' Handle empty table (assign a default value, e.g., 1)
                lastRow = 1
            Else
                lastRow = rs.Fields("MaxID").value
            End If
            
            'Calculate the new ID
            Dim nextID As Long
            nextID = lastRow + 1
        
            'Close the previous recordset, no longer needed and open an empty recordset
            rs.Close
            rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
            
            'Add fields to the recordset before updating it
            rs.AddNew
            rs.Fields("TEC_ID").value = nextID
            rs.Fields("Prof_ID").value = wshAdmin.Range("TEC_Prof_ID")
            rs.Fields("Prof").value = ufSaisieHeures.cmbProfessionnel.value
            rs.Fields("Date").value = CDate(ufSaisieHeures.txtDate.value)
            rs.Fields("Client_ID").value = wshAdmin.Range("TEC_Client_ID")
            rs.Fields("ClientNom").value = ufSaisieHeures.txtClient.value
            rs.Fields("Description").value = ufSaisieHeures.txtActivite.value
            rs.Fields("Heures").value = Format(ufSaisieHeures.txtHeures.value, "#0.00")
            rs.Fields("CommentaireNote").value = ufSaisieHeures.txtCommNote.value
            rs.Fields("EstFacturable").value = ufSaisieHeures.chbFacturable.value
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
                rs.Fields("ClientNom").value = ufSaisieHeures.txtClient.value
                rs.Fields("Description").value = ufSaisieHeures.txtActivite.value
                rs.Fields("Heures").value = Format(ufSaisieHeures.txtHeures.value, "#0.00")
                rs.Fields("CommentaireNote").value = ufSaisieHeures.txtCommNote.value
                rs.Fields("EstFacturable").value = ufSaisieHeures.chbFacturable.value
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
    
    'Free up memory - 2024-02-23
    Set conn = Nothing
    Set rs = Nothing
    
    Application.ScreenUpdating = True

    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        "EffaceLigneDetail() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")

End Sub

Sub Refresh_ListBox_And_Add_Hours() 'Load the listBox with the appropriate records

    Dim timerStart As Double 'Speed tests - 2024-02-20
    timerStart = Timer

    If wshAdmin.Range("TEC_Prof_ID").value = "" Or wshAdmin.Range("TEC_Date").value = "" Then
        GoTo EndOfProcedure
    End If
    
    ufSaisieHeures.txtTotalHeures.value = ""
    
    'Last Row used in first column of result
    Dim lastRow As Long
    lastRow = wshBaseHours.Range("Y99999").End(xlUp).row - 1
    If lastRow = 0 Then Exit Sub
        
    With ufSaisieHeures.lstData
        .ColumnHeads = True
        .ColumnCount = 9
        .ColumnWidths = "28; 26; 51; 130; 180; 35; 80; 32; 83"
        
        If lastRow = 1 Then
            .RowSource = "TEC_Local!Y3:AG3"
        Else
            .RowSource = "TEC_Local!Y3:AG" & lastRow + 1
        End If
    End With

    'Add hours to totalHeures
    Dim nbrRows, i As Integer
    nbrRows = ufSaisieHeures.lstData.ListCount
    Dim totalHeures As Double
    
    If nbrRows > 0 Then
        For i = 0 To nbrRows - 1
            totalHeures = totalHeures + CCur(ufSaisieHeures.lstData.List(i, 5))
        Next
        ufSaisieHeures.txtTotalHeures.value = Format(totalHeures, "#0.00")
    End If

EndOfProcedure:
    ufSaisieHeures.cmdClear.Enabled = False
    ufSaisieHeures.cmdAdd.Enabled = False
    ufSaisieHeures.cmdUpdate.Enabled = False
    ufSaisieHeures.cmdDelete.Enabled = False

    'ufSaisieHeures.txtClient.SetFocus
    
    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        "EffaceLigneDetail() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")
    
End Sub

