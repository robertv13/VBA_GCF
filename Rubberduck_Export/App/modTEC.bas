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

Global Const gAppVersion As String = "v1.1.5"

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

Sub TEC_Import()
    
    Dim startTime As String
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the worksheet
    wshBaseHours.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import TEC from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String
    sourceWorkbook = wshAdmin.Range("SharedFolder").value & Application.PathSeparator & "GCF_BD_Sortie.xlsx" '2023-12-15 @ 19:15

    'Set up source and destination ranges
    Dim sourceRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets("TEC").UsedRange
    'Debug.Print vbNewLine & "Je vais importer toutes les cellules du Range = " & sourceRange.Address & " dans BaseHours!"

    Dim destinationRange As Range
    Set destinationRange = wshBaseHours.Range("A2")

    'Copy data
    sourceRange.Copy destinationRange

    'Close the source workbook
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
    
    TEC_Import '2023-12-15 @ 17:02
    
    With wshBaseHours
        Dim lastRow As Long, LastResultRow As Long, ResultRow As Long
        lastRow = .Range("A999999").End(xlUp).row 'Last BaseHours Row
        If lastRow < 2 Then Exit Sub 'Nothing to filter
        Application.ScreenUpdating = False
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        .Range("A2:Q" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("R2:W3"), _
            CopyToRange:=.Range("Y2:AL2"), _
            Unique:=True
        LastResultRow = .Range("Y999999").End(xlUp).row
        If LastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If LastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
'            .SortFields.Add Key:=wshBaseHours.Range("Z3"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add Key:=wshBaseHours.Range("Y3"), _
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

'************************************************************* AjouteLigneDetail
Sub AjouteLigneDetail()

    If IsDataValid() = False Then Exit Sub
    
    AddOrUpdateTECRecordToDB (0) 'Write to external XLSX file - 2023-12-15 @ 17:09

    'Empty the fields after saving
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
    
    'Reset buttons
    With frmSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdUpdate.Enabled = False
    End With
    
    frmSaisieHeures.txtClient.SetFocus
    
End Sub

Sub ModifieLigneDetail()

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

Sub EffaceLigneDetail()

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
    Set sh = ThisWorkbook.Sheets("HeuresBase")
    
    Dim selectedRow As Long
    'Debug.Print "Le ID du record à DÉTRUIRE, selon Admin est '" & wshAdmin.Range("TEC_Current_ID").value & "'"
    selectedRow = -wshAdmin.Range("TEC_Current_ID").value
    'Debug.Print "Le ID du record à DÉTRUIRE est '" & selectedRow & "'"
        
    AddOrUpdateTECRecordToDB (selectedRow) 'Write to external XLSX file - 2023-12-15 @ 13:33
    
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
    
    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    frmSaisieHeures.txtClient.SetFocus

End Sub


Sub AddOrUpdateTECRecordToDB(r As Long) '2023-12-15 @ 13:33
    Dim FullFileName As String
    Dim SheetName As String
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim MaxID As Long
    Dim lastRow As Long
    Dim nextID As Long
    
    Application.ScreenUpdating = False
    
    'Debug.Print "Dans AddOrUpdateTECRecordToDB, r vaut " & r
    
    FullFileName = wshAdmin.Range("SharedFolder").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    SheetName = "TEC"
    
    'Initialize connection, connection string & open the connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Set rs = CreateObject("ADODB.Recordset")

    'If r is negative, soft delete the record
    If r < 0 Then
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
            MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut être trouvé!", vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    Else
        'If r is 0, add a new record; otherwise, update an existing record
        If r = 0 Then
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
        Else
            'If r is not 0, update an existing record (Only fields that can be different)
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
                MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut être trouvé!", vbExclamation
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

Sub RefreshListBoxAndAddHours()

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

