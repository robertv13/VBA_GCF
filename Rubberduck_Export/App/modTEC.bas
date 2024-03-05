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

Global Const gAppVersion As String = "v2.8" '2024-03-04 @ 10:53

Sub TEC_Ajoute_Ligne() 'Add an entry to DB

    Dim timerStart As Double: timerStart = Timer

    If IsDataValid() = False Then Exit Sub
    
    'Get the Client_ID
    wshAdmin.Range("TEC_Client_ID").value = GetID_From_Client_Name(ufSaisieHeures.txtClient.value)
    
    Call Add_Or_Update_TEC_Record_To_DB(0) 'Write to external XLSX file - 2023-12-23 @ 07:03
    Call Add_Or_Update_TEC_Record_Local(0) 'Write to local worksheet - 2024-02-25 @ 10:34
    
    'Clear the fields after saving
    With ufSaisieHeures
        .txtTEC_ID.value = 0
        .txtClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .chbFacturable = True
    End With

    Call TEC_AdvancedFilter_And_Sort
    Call Refresh_ListBox_And_Add_Hours
    
    'Reset command buttons
    Call Buttons_Enabled_True_Or_False(False, False, False, False)
    
    'Back to client
    ufSaisieHeures.txtClient.SetFocus
    
    Call Output_Timer_Results("TEC_Ajoute_Ligne()", timerStart)

End Sub

Sub TEC_Modifie_Ligne() '2023-12-23 @ 07:04

    Dim timerStart As Double: timerStart = Timer

    If IsDataValid() = False Then Exit Sub

    Call Add_Or_Update_TEC_Record_To_DB(wshAdmin.Range("TEC_Current_ID").value)  'Write to external XLSX file - 2023-12-16 @ 14:10
    Call Add_Or_Update_TEC_Record_Local(wshAdmin.Range("TEC_Current_ID").value)  'Write to local worksheet - 2024-02-25 @ 10:38
 
    'Initialize dynamic variables
    With ufSaisieHeures
        .txtTEC_ID.value = ""
        .cmbProfessionnel.Enabled = True
        .txtDate.Enabled = True
        .txtClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .chbFacturable = True
    End With

    Call TEC_AdvancedFilter_And_Sort
    Call Refresh_ListBox_And_Add_Hours
    
    rmv_state = rmv_modeCreation
    
    ufSaisieHeures.txtClient.SetFocus
    
    Call Output_Timer_Results("TEC_Modifie_Ligne()", timerStart)

End Sub

Sub TEC_Efface_Ligne() '2023-12-23 @ 07:05

    Dim timerStart As Double: timerStart = Timer

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
    
    Dim tecID As Long
    'With a negative ID value, it means to soft delete this record
    tecID = -wshAdmin.Range("TEC_Current_ID").value
    Call Add_Or_Update_TEC_Record_To_DB(tecID)  'Write to external XLSX file - 2023-12-23 @ 07:07
    Call Add_Or_Update_TEC_Record_Local(tecID)  'Write to local worksheet - 2024-02-25 @ 10:40
    
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
    
    Call TEC_AdvancedFilter_And_Sort
    Call Refresh_ListBox_And_Add_Hours
    
    ufSaisieHeures.txtClient.SetFocus

    'Free up memory - 2024-02-23
    Set sh = Nothing

    Call Output_Timer_Results("TEC_Efface_Ligne()", timerStart)

End Sub

Sub TEC_AdvancedFilter_And_Sort() '2024-02-24 @ 09:15
    
    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False

    'Set criteria
    wshBaseHours.Range("R3").value = wshAdmin.Range("TEC_Prof_ID")
    wshBaseHours.Range("S3").value = wshAdmin.Range("TEC_Date")
    wshBaseHours.Range("T3").value = "False"
    
    'ProfID and Date are mandatory to execute this routine
    If wshBaseHours.Range("R3").value = "" Or wshBaseHours.Range("S3").value = "" Then
        Exit Sub
    End If
    
    With wshBaseHours
        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
        lastRow = .Range("A99999").End(xlUp).row 'Last wshBaseHours Used Row
        If lastRow < 3 Then Exit Sub 'Nothing to filter
        
        'Advanced Filter applied to BaseHours (Prof, Date and isDetruit)
        .Range("A2:P" & lastRow).AdvancedFilter _
            xlFilterCopy, _
            CriteriaRange:=.Range("R2:T3"), _
            CopyToRange:=.Range("Y2:AL2"), _
            Unique:=False
        
        lastResultRow = .Range("Y99999").End(xlUp).row
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort 'Sort - Date / Prof / TEC_ID
            .SortFields.clear
            .SortFields.add Key:=wshBaseHours.Range("AA3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.add Key:=wshBaseHours.Range("Z3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Prof_ID
            .SortFields.add Key:=wshBaseHours.Range("Y3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Tec_ID
            .SetRange wshBaseHours.Range("Y3:AL" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With

No_Sort_Required:
    End With
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("TEC_AdvancedFilter_And_Sort()", timerStart)

End Sub

Sub TEC_Efface_Formulaire() 'Clear all fields on the userForm

    Dim timerStart4 As Double: timerStart4 = Timer

    'Empty the dynamic fields after reseting the form
    With ufSaisieHeures
        .txtTEC_ID.value = "" '2024-03-01 @ 09:56
        .txtClient.value = ""
        wshAdmin.Range("TEC_Client_ID").value = 0
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .cmbProfessionnel.Enabled = True
        .txtDate.Enabled = True
    End With
    
    Call TEC_AdvancedFilter_And_Sort
    Call Refresh_ListBox_And_Add_Hours
    
    Call Buttons_Enabled_True_Or_False(False, False, False, False)
        
    ufSaisieHeures.txtClient.SetFocus
    
    Call Output_Timer_Results("TEC_Efface_Formulaire()", timerStart4)

End Sub

Sub Add_Or_Update_TEC_Record_To_DB(tecID As Long) 'Write -OR- Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer

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

    If tecID < 0 Then 'Soft delete a record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & sheetName & "$] WHERE TEC_ID=" & Abs(tecID), conn, 2, 3
        If Not rs.EOF Then
            'Update the "IsDeleted" field to mark the record as deleted
            rs.Fields("DateSaisie").value = Now
            rs.Fields("EstDetruit").value = True
            rs.Fields("VersionApp").value = gAppVersion
            rs.update
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TEC_ID '" & tecID & "' ne peut être trouvé!", _
                vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    Else
        'If r is 0, add a new record; otherwise, update an existing record
        If tecID = 0 Then 'Add a record
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
            wshAdmin.Range("TEC_Current_ID").value = nextID
        
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
            rs.Fields("DateFacture").value = Null
            rs.Fields("EstDetruit").value = False
            rs.Fields("VersionApp").value = gAppVersion
            rs.Fields("NoFacture").value = ""
        Else 'Update an existing record
            'Open the recordset for the specified ID
            rs.Open "SELECT * FROM [" & sheetName & "$] WHERE TEC_ID=" & tecID, conn, 2, 3
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
                MsgBox "L'enregistrement avec le TEC_ID '" & tecID & "' ne peut être trouvé!", vbExclamation
                rs.Close
                conn.Close
                Exit Sub
            End If
        End If
    End If
    'Update the recordset (create the record)
    rs.update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Free up memory - 2024-02-23
    Set conn = Nothing
    Set rs = Nothing
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("Add_Or_Update_TEC_Record_To_DB()", timerStart)

End Sub

Sub Add_Or_Update_TEC_Record_Local(tecID As Long) 'Write -OR- Update a record to local worksheet
    
    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    'What is the row number of this TEC_ID ?
    Dim lastUsedRow As Long
    
    Dim hoursValue As Double '2024-03-01 @ 05:40
    hoursValue = CDbl(ufSaisieHeures.txtHeures.value)
    
    If tecID = 0 Then 'Add a new record
        'Get the next available row in TEC_Local
        Dim nextRowNumber As Long
        nextRowNumber = wshBaseHours.Range("A9999").End(xlUp).row + 1
        With wshBaseHours
            .Range("A" & nextRowNumber).value = wshAdmin.Range("TEC_Current_ID").value
            .Range("B" & nextRowNumber).value = wshAdmin.Range("TEC_Prof_ID").value
            .Range("C" & nextRowNumber).value = ufSaisieHeures.cmbProfessionnel.value
            .Range("D" & nextRowNumber).value = CDate(ufSaisieHeures.txtDate.value)
            .Range("E" & nextRowNumber).value = wshAdmin.Range("TEC_Client_ID").value
            .Range("F" & nextRowNumber).value = ufSaisieHeures.txtClient.value
            .Range("G" & nextRowNumber).value = ufSaisieHeures.txtActivite.value
            .Range("H" & nextRowNumber).value = hoursValue
            .Range("I" & nextRowNumber).value = ufSaisieHeures.txtCommNote.value
            .Range("J" & nextRowNumber).value = ufSaisieHeures.chbFacturable.value
            .Range("K" & nextRowNumber).value = Now()
            .Range("L" & nextRowNumber).value = False
            .Range("M" & nextRowNumber).value = ""
            .Range("N" & nextRowNumber).value = False
            .Range("O" & nextRowNumber).value = gAppVersion
            .Range("P" & nextRowNumber).value = ""
        End With
    Else
        'What is the row number for the TEC_ID
        Dim lookupRange As Range, rowToBeUpdated As Long
        lastUsedRow = wshBaseHours.Range("A99999").End(xlUp).row
        Set lookupRange = wshBaseHours.Range("A3:A" & lastUsedRow)
        rowToBeUpdated = Get_TEC_Row_Number_By_TEC_ID(Abs(tecID), lookupRange)
        If rowToBeUpdated = 0 Then
            'Handle the case where the specified TecID is not found !!
            MsgBox "L'enregistrement avec le TEC_ID '" & tecID & "' ne peut être trouvé!", _
                vbExclamation
            Exit Sub
        End If

        If tecID > 0 Then 'Modify the record
            With wshBaseHours
                .Range("E" & rowToBeUpdated).value = wshAdmin.Range("TEC_Client_ID").value
                .Range("F" & rowToBeUpdated).value = ufSaisieHeures.txtClient.value
                .Range("G" & rowToBeUpdated).value = ufSaisieHeures.txtActivite.value
                .Range("H" & rowToBeUpdated).value = hoursValue
                .Range("I" & rowToBeUpdated).value = ufSaisieHeures.txtCommNote.value
                .Range("J" & rowToBeUpdated).value = ufSaisieHeures.chbFacturable.value
                .Range("K" & rowToBeUpdated).value = Now()
                .Range("L" & rowToBeUpdated).value = False
                .Range("M" & rowToBeUpdated).value = ""
                .Range("N" & rowToBeUpdated).value = False
                .Range("O" & rowToBeUpdated).value = gAppVersion
                .Range("P" & rowToBeUpdated).value = ""
            End With
        Else 'Soft delete the record
            wshBaseHours.Range("K" & rowToBeUpdated).value = Now()
            wshBaseHours.Range("N" & rowToBeUpdated).value = True
            wshBaseHours.Range("O" & rowToBeUpdated).value = gAppVersion
        End If
    End If
    
    'Free up memory - 2024-02-23
    Set lookupRange = Nothing
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("Add_Or_Update_TEC_Record_Local()", timerStart)

End Sub

Sub Refresh_ListBox_And_Add_Hours() 'Load the listBox with the appropriate records

    Dim timerStart As Double: timerStart = Timer

    If wshAdmin.Range("TEC_Prof_ID").value = "" Or wshAdmin.Range("TEC_Date").value = "" Then
        GoTo EndOfProcedure
    End If
    
    ufSaisieHeures.txtTotalHeures.value = ""
    
    'Last Row used in first column of result
    Dim lastRow As Long
    lastRow = wshBaseHours.Range("Y999").End(xlUp).row
    If lastRow < 3 Then Exit Sub
        
    With ufSaisieHeures.ListBox2
        .ColumnHeads = True
        .ColumnCount = 10
        .ColumnWidths = "30; 26; 52; 130; 200; 35; 80; 38; 83"
        .RowSource = "TEC_Local!Y3:AG" & lastRow
    End With

    'Add hours to totalHeures
    Dim nbrRows, i As Integer
    nbrRows = ufSaisieHeures.ListBox2.ListCount
    Dim totalHeures As Double
    
    If nbrRows > 0 Then
        For i = 0 To nbrRows - 1
            totalHeures = totalHeures + CCur(ufSaisieHeures.ListBox2.List(i, 5))
        Next
        ufSaisieHeures.txtTotalHeures.value = Format(totalHeures, "#0.00")
    End If

EndOfProcedure:

    Call Buttons_Enabled_True_Or_False(False, False, False, False)

    ufSaisieHeures.txtClient.SetFocus
    
    Call Output_Timer_Results("Refresh_ListBox_And_Add_Hours()", timerStart)
    
End Sub

