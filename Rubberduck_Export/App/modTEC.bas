Attribute VB_Name = "modTEC"
Option Explicit

Global Const rmv_modeInitial As Long = 1
Global Const rmv_modeCreation As Long = 2
Global Const rmv_modeAffichage As Long = 3
Global Const rmv_modeModification As Long = 4

Global rmv_state As Long

Global savedClient As String
Global savedActivite As String
Global savedHeures As String
Global savedFacturable As String
Global savedCommNote As String

Sub TEC_Ajoute_Ligne() 'Add an entry to DB

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_Ajoute_Ligne", 0)

    If Fn_TEC_Is_Data_Valid() = True Then
        'Get the Client_ID
        wshAdmin.Range("TEC_Client_ID").value = Fn_GetID_From_Client_Name(ufSaisieHeures.txtClient.value)
        
        Call TEC_Record_Add_Or_Update_To_DB(0) 'Write to MASTER.xlsx file - 2023-12-23 @ 07:03
        Call TEC_Record_Add_Or_Update_Locally(0) 'Write to local worksheet - 2024-02-25 @ 10:34
        
        'Clear the fields after saving
        With ufSaisieHeures
            .txtTEC_ID.value = ""
            .txtClient.value = ""
            .txtActivite.value = ""
            .txtHeures.value = ""
            .txtCommNote.value = ""
            .chbFacturable = True
        End With
        
        Call TEC_AdvancedFilter_And_Sort
        Call TEC_Refresh_ListBox_And_Add_Hours
        
        Call TEC_TdB_Update_All
        
        'Reset command buttons
        Call Buttons_Enabled_True_Or_False(False, False, False, False)
        
        Call SetNumLockOn '2024-08-26 @ 09:54
        
        'Back to client
        ufSaisieHeures.txtClient.SetFocus
    End If
    
    Call Log_Record("modTEC:TEC_Ajoute_Ligne()", startTime)

End Sub

Sub TEC_Modifie_Ligne() '2023-12-23 @ 07:04

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_Modifie_Ligne", 0)

    If Fn_TEC_Is_Data_Valid() = False Then Exit Sub

    Call TEC_Record_Add_Or_Update_To_DB(wshAdmin.Range("TEC_Current_ID").value)  'Write to external XLSX file - 2023-12-16 @ 14:10
    Call TEC_Record_Add_Or_Update_Locally(wshAdmin.Range("TEC_Current_ID").value)  'Write to local worksheet - 2024-02-25 @ 10:38
 
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
    Call TEC_Refresh_ListBox_And_Add_Hours
    
    rmv_state = rmv_modeCreation
    
    ufSaisieHeures.txtClient.SetFocus
    
    Call Log_Record("modTEC:TEC_Modifie_Ligne()", startTime)

End Sub

Sub TEC_Efface_Ligne() '2023-12-23 @ 07:05

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_Efface_Ligne", 0)

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox _
        Prompt:="Vous devez choisir un enregistrement � D�TRUIRE !", _
        Buttons:=vbCritical
        GoTo Clean_Exit
    End If
    
    Dim answerYesNo As Long
    answerYesNo = MsgBox("�tes-vous certain de vouloir D�TRUIRE cet enregistrement ? ", _
                         vbYesNo + vbQuestion, "Confirmation de DESTRUCTION")
    If answerYesNo = vbNo Then
        MsgBox _
        Prompt:="Cet enregistrement ne sera PAS d�truit ! ", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        GoTo Clean_Exit
    End If
    
    Dim Sh As Worksheet: Set Sh = wshTEC_Local
    
    Dim TECID As Long
    'With a negative ID value, it means to soft delete this record
    TECID = -wshAdmin.Range("TEC_Current_ID").value
    Call TEC_Record_Add_Or_Update_To_DB(TECID)  'Write to external XLSX file - 2023-12-23 @ 07:07
    Call TEC_Record_Add_Or_Update_Locally(TECID)  'Write to local worksheet - 2024-02-25 @ 10:40
    
    'Empty the dynamic fields after deleting
    With ufSaisieHeures
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
        
    ufSaisieHeures.cmbProfessionnel.Enabled = True
    ufSaisieHeures.txtDate.Enabled = True
    rmv_state = rmv_modeCreation
    
    Call TEC_AdvancedFilter_And_Sort
    Call TEC_Refresh_ListBox_And_Add_Hours
    
Clean_Exit:

    ufSaisieHeures.txtTEC_ID.value = ""
    ufSaisieHeures.txtClient.SetFocus

    'Cleaning memory - 2024-07-01 @ 09:34
    Set Sh = Nothing

    Call Log_Record("modTEC:TEC_Efface_Ligne()", startTime)

End Sub

Sub TEC_AdvancedFilter_And_Sort() '2024-02-24 @ 09:15
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_AdvancedFilter_And_Sort", 0)

    Application.ScreenUpdating = False

    'ProfID and Date are mandatory to execute this routine
    If wshAdmin.Range("TEC_Prof_ID") = "" Or wshAdmin.Range("TEC_Date") = "" Then
        Exit Sub
    End If
    
    'Set criteria in worksheet
    wshTEC_Local.Range("R3").value = wshAdmin.Range("TEC_Prof_ID")
    wshTEC_Local.Range("S3").value = wshAdmin.Range("TEC_Date")
    wshTEC_Local.Range("T3").value = "FAUX"
    
    With wshTEC_Local
        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
        lastRow = .Range("A99999").End(xlUp).Row 'Last wshTEC_Local Used Row
        If lastRow < 3 Then Exit Sub 'Nothing to filter
        
        'Data Source
        Dim sRng As Range: Set sRng = .Range("A2:P" & lastRow)
        
         'Criteria
        Dim cRng As Range: Set cRng = .Range("R2:T3")
        
        'Destination
        Dim dRng As Range: Set dRng = .Range("V2:AI2")
        
        'Advanced Filter applied to BaseHours (Prof, Date and isDetruit)
        sRng.AdvancedFilter xlFilterCopy, cRng, dRng, False
        
        lastResultRow = .Range("V99999").End(xlUp).Row
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort 'Sort - Date / Prof / TEC_ID
            .SortFields.clear
            .SortFields.add key:=wshTEC_Local.Range("X3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.add key:=wshTEC_Local.Range("W3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Prof
            .SortFields.add key:=wshTEC_Local.Range("V3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Tec_ID
            .SetRange wshTEC_Local.Range("V3:AI" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With

No_Sort_Required:
    End With
    
    'Suddenly, I have to convert BOOLEAN value to TEXT !!!! - 2024-06-19 @ 14:20
    If lastResultRow > 2 Then
        Set dRng = wshTEC_Local.Range("AC3:AC" & lastResultRow)
        Call ConvertRangeBooleanToText(dRng)
        Set dRng = wshTEC_Local.Range("AE3:AE" & lastResultRow)
        Call ConvertRangeBooleanToText(dRng)
        Set dRng = wshTEC_Local.Range("AG3:AG" & lastResultRow)
        Call ConvertRangeBooleanToText(dRng)
    End If
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set cRng = Nothing
    Set dRng = Nothing
    Set sRng = Nothing
    
    Call Log_Record("modTEC:TEC_AdvancedFilter_And_Sort()", startTime)

End Sub

Sub Test_Advanced_Filter_1() '2024-06-19 @ 16:20
    
    Application.ScreenUpdating = False

    With wshTEC_Local
        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
        lastRow = .Range("A99999").End(xlUp).Row 'Last wshTEC_Local Used Row
        If lastRow < 3 Then Exit Sub 'Nothing to filter
        
        'Data Source
        Dim sRng As Range: Set sRng = .Range("A2:P" & lastRow)
        .Range("S10").value = sRng.Address
        
        'Criteria
        Dim cRng As Range: Set cRng = .Range("R2:T3")
        .Range("S11").value = cRng.Address
        
        'Destination
        Dim dRng As Range: Set dRng = .Range("V2:AI2")
        .Range("S12").value = dRng.Address
        
        'Advanced Filter applied to BaseHours (Prof, Date and isDetruit)
        sRng.AdvancedFilter xlFilterCopy, cRng, dRng, False
        
        lastResultRow = .Range("V99999").End(xlUp).Row
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort 'Sort - Date / Prof / TEC_ID
            .SortFields.clear
            .SortFields.add key:=wshTEC_Local.Range("X3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.add key:=wshTEC_Local.Range("W3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Prof
            .SortFields.add key:=wshTEC_Local.Range("V3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Tec_ID
            .SetRange wshTEC_Local.Range("V3:AI" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With

No_Sort_Required:
    End With
    
    'Suddenly, I have to convert BOOLEAN value to TEXT !!!! - 2024-06-19 @ 14:20
    If lastResultRow > 2 Then
        Set dRng = wshTEC_Local.Range("AC3:AC" & lastResultRow)
        Call ConvertRangeBooleanToText(dRng)
        Set dRng = wshTEC_Local.Range("AE3:AE" & lastResultRow)
        Call ConvertRangeBooleanToText(dRng)
        Set dRng = wshTEC_Local.Range("AG3:AG" & lastResultRow)
        Call ConvertRangeBooleanToText(dRng)
    End If
    
    wshTEC_Local.Range("S13").value = lastResultRow - 2 & " rows"
    wshTEC_Local.Range("S14").value = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set cRng = Nothing
    Set dRng = Nothing
    Set sRng = Nothing
    
End Sub

Sub TEC_Efface_Formulaire() 'Clear all fields on the userForm

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_Efface_Formulaire", 0)

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
    Call TEC_Refresh_ListBox_And_Add_Hours
    
    Call Buttons_Enabled_True_Or_False(False, False, False, False)
        
    ufSaisieHeures.txtClient.SetFocus
    
    Call Log_Record("modTEC:TEC_Efface_Formulaire()", startTime)

End Sub

Sub TEC_Record_Add_Or_Update_To_DB(TECID As Long) 'Write -OR- Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_Record_Add_Or_Update_To_DB", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "TEC_Local"
    
    On Error GoTo ErrorHandler
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    Dim strConnection As String
    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
        destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    conn.Open strConnection
    
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim saveLogTEC_ID As Long
    saveLogTEC_ID = TECID
    
    If TECID < 0 Then 'Soft delete a record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & Abs(TECID), conn, 2, 3
        saveLogTEC_ID = TECID
        If Not rs.EOF Then
            'Update the "IsDeleted" field to mark the record as deleted
            rs.Fields("DateSaisie").value = CDate(Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            rs.Fields("EstDetruit").value = ConvertValueBooleanToText(True)
            rs.Fields("VersionApp").value = ThisWorkbook.name
            rs.update
            Call Log_Saisie_Heures("D", CStr(saveLogTEC_ID)) '2024-09-02 @ 10:35
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TEC_ID '" & TECID & "' ne peut �tre trouv�!", _
                vbExclamation
            rs.Close
            conn.Close
            Call Log_Saisie_Heures("?", CStr(saveLogTEC_ID)) '2024-09-02 @ 10:35
            Exit Sub
        End If
    Else
        'If r is 0, add a new record; otherwise, update an existing record
        If TECID = 0 Then 'Add a record
        'SQL select command to find the next available ID
            Dim strSQL As String, MaxID As Long
            strSQL = "SELECT MAX(TEC_ID) AS MaxID FROM [" & destinationTab & "$]"
        
            'Open recordset to find out the MaxID
            rs.Open strSQL, conn
            
            'Get the last used row
            Dim lastRow As Long
            If IsNull(rs.Fields("MaxID").value) Then
                'Handle empty table (assign a default value, e.g., 0)
                lastRow = 0
            Else
                lastRow = rs.Fields("MaxID").value
            End If
            
            'Calculate the new ID
            Dim nextID As Long
            nextID = lastRow + 1
            wshAdmin.Range("TEC_Current_ID").value = nextID
            saveLogTEC_ID = nextID
        
            'Close the previous recordset, no longer needed and open an empty recordset
            rs.Close
            rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
            
            'Add fields to the recordset before updating it
            rs.AddNew
            rs.Fields("TEC_ID").value = nextID
            rs.Fields("Prof_ID").value = wshAdmin.Range("TEC_Prof_ID")
            rs.Fields("Prof").value = ufSaisieHeures.cmbProfessionnel.value
            rs.Fields("Date").value = CDate(Format$(ufSaisieHeures.txtDate.value, "dd/mm/yyyy"))
            rs.Fields("Client_ID").value = wshAdmin.Range("TEC_Client_ID")
            rs.Fields("ClientNom").value = ufSaisieHeures.txtClient.value
            If Len(ufSaisieHeures.txtActivite.value) > 255 Then
                ufSaisieHeures.txtActivite.value = Left(ufSaisieHeures.txtActivite.value, 255)
            End If
            rs.Fields("Description").value = ufSaisieHeures.txtActivite.value
            rs.Fields("Heures").value = Format$(ufSaisieHeures.txtHeures.value, "#0.00")
            rs.Fields("CommentaireNote").value = ufSaisieHeures.txtCommNote.value
            rs.Fields("EstFacturable").value = ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value)
            rs.Fields("DateSaisie").value = CDate(Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            rs.Fields("EstFacturee").value = ConvertValueBooleanToText(False)
            rs.Fields("DateFacturee").value = Null
            rs.Fields("EstDetruit").value = ConvertValueBooleanToText(False)
            rs.Fields("VersionApp").value = ThisWorkbook.name
            rs.Fields("NoFacture").value = ""
            'Nouveau log - 2024-09-02 @ 10:40
            Call Log_Saisie_Heures("Add", saveLogTEC_ID & "|" & _
                ufSaisieHeures.cmbProfessionnel.value & "|" & _
                CDate(Format$(ufSaisieHeures.txtDate.value, "dd/mm/yyyy")) & "|" & _
                wshAdmin.Range("TEC_Client_ID") & "|" & _
                ufSaisieHeures.txtClient.value & "|" & _
                ufSaisieHeures.txtActivite.value & "|" & _
                Format$(ufSaisieHeures.txtHeures.value, "#0.00") & "|" & _
                ufSaisieHeures.txtCommNote.value & "|" & _
                ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value))
        Else 'Update an existing record
            'Open the recordset for the specified ID
            rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & TECID, conn, 2, 3
            If Not rs.EOF Then
                'Update fields for the existing record
                rs.Fields("Client_ID").value = wshAdmin.Range("TEC_Client_ID")
                rs.Fields("ClientNom").value = ufSaisieHeures.txtClient.value
                rs.Fields("Description").value = ufSaisieHeures.txtActivite.value
                rs.Fields("Heures").value = Format$(ufSaisieHeures.txtHeures.value, "#0.00")
                rs.Fields("CommentaireNote").value = ufSaisieHeures.txtCommNote.value
                rs.Fields("EstFacturable").value = ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value)
                rs.Fields("DateSaisie").value = CDate(Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
                rs.Fields("VersionApp").value = ThisWorkbook.name
                'Nouveau log - 2024-09-02 @ 10:40
                Call Log_Saisie_Heures("Update", saveLogTEC_ID & "|" & _
                    ufSaisieHeures.cmbProfessionnel.value & "|" & _
                    CDate(Format$(ufSaisieHeures.txtDate.value, "dd/mm/yyyy")) & "|" & _
                    wshAdmin.Range("TEC_Client_ID") & "|" & _
                    ufSaisieHeures.txtClient.value & "|" & _
                    ufSaisieHeures.txtActivite.value & "|" & _
                    Format$(ufSaisieHeures.txtHeures.value, "#0.00") & "|" & _
                    ufSaisieHeures.txtCommNote.value & "|" & _
                    ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value))
            Else
                'Handle the case where the specified ID is not found
                MsgBox "L'enregistrement avec le TEC_ID '" & TECID & "' ne peut �tre trouv�!", vbExclamation
                Call Log_Saisie_Heures("??", CStr(saveLogTEC_ID)) '2024-09-02 @ 10:35
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
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modTEC:TEC_Record_Add_Or_Update_To_DB()", startTime)

    Exit Sub
    
ErrorHandler:

     'Si une erreur survient, cela signifie que le fichier est en lecture seule
    MsgBox "Le fichier 'MASTER' est en lecture seule" & vbNewLine & vbNewLine & _
           "ou d�j� ouvert par un autre utilisateur.", vbCritical, "Probl�me MAJEUR"
    On Error GoTo 0
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
    
End Sub

Sub TEC_Record_Add_Or_Update_Locally(TECID As Long) 'Write -OR- Update a record to local worksheet
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_Record_Add_Or_Update_Locally", 0)

    Application.ScreenUpdating = False
    
    'What is the row number of this TEC_ID ?
    Dim lastUsedRow As Long
    
    Dim hoursValue As Double '2024-03-01 @ 05:40
    hoursValue = CDbl(ufSaisieHeures.txtHeures.value)
    
    If TECID = 0 Then 'Add a new record
        'Get the next available row in TEC_Local
        Dim nextRowNumber As Long
        nextRowNumber = wshTEC_Local.Range("A9999").End(xlUp).Row + 1
        With wshTEC_Local
            .Range("A" & nextRowNumber).value = wshAdmin.Range("TEC_Current_ID").value
            .Range("B" & nextRowNumber).value = wshAdmin.Range("TEC_Prof_ID").value
            .Range("C" & nextRowNumber).value = ufSaisieHeures.cmbProfessionnel.value
            .Range("D" & nextRowNumber).value = CDate(Format$(ufSaisieHeures.txtDate.value, "dd/mm/yyyy"))
            .Range("E" & nextRowNumber).value = wshAdmin.Range("TEC_Client_ID").value
            .Range("F" & nextRowNumber).value = ufSaisieHeures.txtClient.value
            .Range("G" & nextRowNumber).value = ufSaisieHeures.txtActivite.value
            .Range("H" & nextRowNumber).value = hoursValue
            .Range("I" & nextRowNumber).value = ufSaisieHeures.txtCommNote.value
            .Range("J" & nextRowNumber).value = ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value)
            .Range("K" & nextRowNumber).value = CDate(Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            .Range("L" & nextRowNumber).value = ConvertValueBooleanToText(False)
            .Range("M" & nextRowNumber).value = ""
            .Range("N" & nextRowNumber).value = ConvertValueBooleanToText(False)
            .Range("O" & nextRowNumber).value = ThisWorkbook.name
            .Range("P" & nextRowNumber).value = ""
        End With
    Else
        'What is the row number for the TEC_ID
        lastUsedRow = wshTEC_Local.Range("A99999").End(xlUp).Row
        Dim lookupRange As Range:  Set lookupRange = wshTEC_Local.Range("A3:A" & lastUsedRow)
        Dim rowToBeUpdated As Long
        rowToBeUpdated = Fn_Find_Row_Number_TEC_ID(Abs(TECID), lookupRange)
        If rowToBeUpdated < 1 Then
            'Handle the case where the specified TecID is not found !!
            MsgBox "L'enregistrement avec le TEC_ID '" & TECID & "' ne peut �tre trouv�!", _
                vbExclamation
            Exit Sub
        End If

        If TECID > 0 Then 'Modify the record
            With wshTEC_Local
                .Range("E" & rowToBeUpdated).value = wshAdmin.Range("TEC_Client_ID").value
                .Range("F" & rowToBeUpdated).value = ufSaisieHeures.txtClient.value
                .Range("G" & rowToBeUpdated).value = ufSaisieHeures.txtActivite.value
                .Range("H" & rowToBeUpdated).value = hoursValue
                .Range("I" & rowToBeUpdated).value = ufSaisieHeures.txtCommNote.value
                .Range("J" & rowToBeUpdated).value = ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value)
                .Range("K" & rowToBeUpdated).value = CDate(Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
                .Range("L" & rowToBeUpdated).value = ConvertValueBooleanToText(False)
                .Range("M" & rowToBeUpdated).value = ""
                .Range("N" & rowToBeUpdated).value = ConvertValueBooleanToText(False)
                .Range("O" & rowToBeUpdated).value = ThisWorkbook.name
                .Range("P" & rowToBeUpdated).value = ""
            End With
        Else 'Soft delete the record
            wshTEC_Local.Range("K" & rowToBeUpdated).value = CDate(Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            wshTEC_Local.Range("N" & rowToBeUpdated).value = ConvertValueBooleanToText(True)
            wshTEC_Local.Range("O" & rowToBeUpdated).value = ThisWorkbook.name
        End If
    End If
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set lookupRange = Nothing
    
    Call Log_Record("modTEC:TEC_Record_Add_Or_Update_Locally()", startTime)

End Sub

Sub TEC_Refresh_ListBox_And_Add_Hours() 'Load the listBox with the appropriate records

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_Refresh_ListBox_And_Add_Hours", 0)

    If wshAdmin.Range("TEC_Prof_ID").value = "" Or wshAdmin.Range("TEC_Date").value = "" Then
        GoTo EndOfProcedure
    End If
    
    ufSaisieHeures.txtTotalHeures.value = ""
    ufSaisieHeures.lsbHresJour.RowSource = ""
    ufSaisieHeures.lsbHresJour.clear '2024-08-10 @ 05:59
    
    'Last Row used in first column of result
    Dim lastRow As Long
    lastRow = wshTEC_Local.Range("V999").End(xlUp).Row
    If lastRow < 3 Then Exit Sub
        
    With ufSaisieHeures.lsbHresJour
        .ColumnHeads = False
        .ColumnCount = 9
        .ColumnWidths = "30; 24; 54; 155; 240; 35; 90; 32; 90"
'        .RowSource = wshTEC_Local.name & "!V3:AI" & lastRow '2024-08-11 @ 12:50
    End With
    
    'Manually add to listBox (.RowSource DOES NOT WORK!!!)
    Dim rng As Range
    Set rng = wshTEC_Local.Range("V3:AI" & lastRow)
    Debug.Print rng.Address
     
    Dim i As Long, j As Long
    Dim totalHeures As Double
    Application.ScreenUpdating = True
    For i = 1 To rng.rows.count
        ufSaisieHeures.lsbHresJour.AddItem rng.Cells(i, 1).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 1) = rng.Cells(i, 2).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 2) = Format$(rng.Cells(i, 3).value, "dd/mm/yyyy")
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 3) = rng.Cells(i, 4).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 4) = rng.Cells(i, 5).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 5) = Format$(rng.Cells(i, 6).value, "#,##0.00")
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 6) = rng.Cells(i, 7).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 7) = rng.Cells(i, 8).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 8) = Format$(rng.Cells(i, 9).value, "dd/mm/yyyy hh:nn:ss")
        totalHeures = totalHeures + CCur(rng.Cells(i, 6).value)
    Next i
         
    ufSaisieHeures.Repaint
    
    ufSaisieHeures.txtTotalHeures.value = Format$(totalHeures, "#0.00")
    
    DoEvents '2024-08-12 @ 10:31
    
    Application.ScreenUpdating = True

EndOfProcedure:

    Call Buttons_Enabled_True_Or_False(False, False, False, False)

    ufSaisieHeures.txtClient.SetFocus
    
    Call Log_Record("modTEC:TEC_Refresh_ListBox_And_Add_Hours()", startTime)
    
End Sub

Sub TEC_TdB_Push_TEC_Local_To_DB_Data()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_TdB_Push_TEC_Local_To_DB_Data", 0)

    Dim wsFrom As Worksheet: Set wsFrom = wshTEC_Local
    
    Dim lastUsedRow As Long
    lastUsedRow = wshTEC_Local.Range("A99999").End(xlUp).Row
    
    Dim arr() As Variant
    ReDim arr(1 To lastUsedRow - 2, 1 To 8) '2 rows of Heading
    
    Dim i As Long
    For i = 3 To lastUsedRow
        With wsFrom
            arr(i - 2, 1) = .Range("A" & i).value 'TEC_ID
            arr(i - 2, 2) = .Range("C" & i).value 'Prof
            arr(i - 2, 3) = .Range("D" & i).value 'Date
            arr(i - 2, 4) = .Range("F" & i).value 'Client's Name
            arr(i - 2, 5) = .Range("H" & i).value 'Hours
            arr(i - 2, 6) = .Range("J" & i).value 'isBillable
            arr(i - 2, 7) = .Range("L" & i).value 'isInvoiced
            arr(i - 2, 8) = .Range("N" & i).value 'isDeleted
        End With
    Next i

    Dim rngTo As Range: Set rngTo = wshTEC_TDB_Data.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2))
    rngTo.value = arr
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rngTo = Nothing
    Set wsFrom = Nothing
    
    Call Log_Record("modTEC:TEC_TdB_Push_TEC_Local_To_DB_Data()", startTime)

End Sub

Sub TEC_TdB_Update_All()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_TdB_Update_All", 0)
    
    Call TEC_TdB_Push_TEC_Local_To_DB_Data
    Call TEC_TdB_Refresh_All_Pivot_Tables
    
    Call Log_Record("modTEC:TEC_TdB_Update_All()", startTime)

End Sub

Sub TEC_TdB_Refresh_All_Pivot_Tables()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC:TEC_TdB_Refresh_All_Pivot_Tables", 0)
    
    Dim pt As PivotTable
    For Each pt In wshTEC_TDB_PivotTable.PivotTables
        pt.RefreshTable
    Next pt

    Call Log_Record("modTEC:TEC_TdB_Refresh_All_Pivot_Tables()", startTime)

    'Cleaning memory - 2024-07-01 @ 09:34
    Set pt = Nothing
    
End Sub

Sub TEC_Advanced_Filter_2() 'Advanced Filter for TEC records - 2024-06-19 @ 12:41
    
    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    With ws
        Dim lastUsedRow As Long
        lastUsedRow = .Range("A99999").End(xlUp).Row
        Dim sRng As Range: Set sRng = .Range("A2:P" & lastUsedRow)
        .Range("AL10").value = sRng.Address & " - " & _
            .Range("A2:P" & lastUsedRow).rows.count & " rows, " & _
            .Range("A2:P" & lastUsedRow).columns.count & " columns"
        
        Dim filterDate As Date
        filterDate = DateValue("02/09/2024")
        .Range("AL3").value = "<=" & filterDate
        .Range("AL3").NumberFormat = "dd/mm/yyyy"
        
'        .Range("AL3").value = "<=" & Format(DateSerial(2024, 8, 23), "DD/mm/yyyy")
        Dim cRng As Range: Set cRng = .Range("AK2:AO3")
        .Range("AL11").value = cRng.Address & " - " & .Range("AK2:AO3").columns.count & " columns"
        
        lastUsedRow = .Range("AQ99999").End(xlUp).Row
        Dim dRng As Range: Set dRng = .Range("AQ2:BE" & lastUsedRow)
        dRng.Offset(1, 0).ClearContents
        .Range("AL12").value = dRng.Address & " - " & .Range("AQ2:BE" & lastUsedRow).columns.count & " columns"
        
        On Error GoTo ErrorHandler
        sRng.AdvancedFilter action:=xlFilterCopy, _
                            criteriaRange:=cRng, _
                            CopyToRange:=dRng, _
                            Unique:=False
        On Error GoTo 0
        
        Dim lastResultRow As Long
        lastResultRow = .Range("AQ99999").End(xlUp).Row
            If lastResultRow < 4 Then GoTo No_Sort_Required
            With .Sort
                .SortFields.clear
                .SortFields.add key:=wshTEC_Local.Range("AT3"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Sort Based On Date
                .SortFields.add key:=wshTEC_Local.Range("AR3"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Sort Based On Prof_ID
                .SortFields.add key:=wshTEC_Local.Range("AQ3"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Sort Based On TEC_ID
                .SetRange wshTEC_Local.Range("AQ3:BE" & lastResultRow) 'Set Range
                .Apply 'Apply Sort
             End With
    
No_Sort_Required:
        wshTEC_Local.Range("AL13").value = lastResultRow - 2 & " rows"
        wshTEC_Local.Range("AL14").value = Format$(Now(), "mm/dd/yyyy hh:mm:ss")
    End With
    
Cleaning:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set cRng = Nothing
    Set dRng = Nothing
    Set sRng = Nothing
    Set ws = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite lors de l'application du filtre avanc�.", vbCritical
    On Error GoTo 0
    Resume Cleaning

End Sub

Sub Buttons_Enabled_True_Or_False(clear As Boolean, add As Boolean, _
                                  update As Boolean, delete As Boolean)
    With ufSaisieHeures
        .cmdClear.Enabled = clear
        .cmdAdd.Enabled = add
        .cmdUpdate.Enabled = update
        .cmdDelete.Enabled = delete
    End With

End Sub

Sub MsgBoxInvalidDate() '2024-06-13 @ 12:40

    MsgBox "La date saisie ne peut �tre accept�e tel qu'elle est entr�e." & vbNewLine & vbNewLine & _
           "Elle doit �tre obligatoirement de format:" & vbNewLine & _
           "     'jj', " & vbNewLine & _
           "     'jj-mm' ou " & vbNewLine & _
           "     'jj-mm-aaaa'" & vbNewLine & vbNewLine & _
           "Veuillez saisir la date de nouveau SVP", _
           vbCritical, _
           "La date saisie est INVALIDE"

End Sub


