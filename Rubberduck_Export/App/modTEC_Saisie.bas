Attribute VB_Name = "modTEC_Saisie"
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

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Ajoute_Ligne", 0)

    If Fn_TEC_Is_Data_Valid() = True Then
        Dim Y As Integer, m As Integer, d As Integer
        Dim avant As String, apres As String
        On Error Resume Next
            avant = ufSaisieHeures.txtDate.value
            Y = year(ufSaisieHeures.txtDate.value)
            m = month(ufSaisieHeures.txtDate.value)
            d = day(ufSaisieHeures.txtDate.value)
            If Y = 2024 And m < 9 Then 'Si mois < 9 alors, on prend pour acquis que le jour et le mois sont inversés...
                Dim temp As Integer
                temp = m
                m = d
                d = temp
                Call Log_Saisie_Heures("info     ", "@00045 - AJUSTEMENT (PLUG) --->   y = " & Y & "   m = " & m & "   d = " & d & "   type = " & TypeName(ufSaisieHeures.txtDate.value))
            End If
            ufSaisieHeures.txtDate.value = Format$(DateSerial(Y, m, d), "yyyy-mm-dd")
            Call Log_Saisie_Heures("info     ", "@00048 - ufSaisieHeures.txtDate.value = " & ufSaisieHeures.txtDate.value & "   type = " & TypeName(ufSaisieHeures.txtDate.value) & "   après assignation")
            apres = ufSaisieHeures.txtDate.value
            If apres <> avant Then
                Call Log_Saisie_Heures("adjust   ", "@00051 - La date a été changée pour corriger la date - " & avant & " ---> " & apres)
            End If
        On Error GoTo 0
        
        Call TEC_Record_Add_Or_Update_To_DB(0)
        
        Call TEC_Record_Add_Or_Update_Locally(0)
        
        'Clear the userForm fields after saving
        With ufSaisieHeures
            .txtTEC_ID.value = ""
            .txtClient.value = ""
            .txtClient_ID.value = ""
            .txtActivite.value = ""
            .txtHeures.value = ""
            .txtCommNote.value = ""
            .chbFacturable = True
            .txtSavedHeures.value = ""
        End With
        
        Call TEC_Get_All_TEC_AF
        
        Call TEC_Refresh_ListBox_And_Add_Hours
        
'        Call TEC_TdB_Update_All
        
        'Reset command buttons
        Call ActiverButtonsVraiOuFaux(False, False, False, False)
        
        'Back to client
        ufSaisieHeures.txtClient.SetFocus
    End If
    
    Call Log_Record("modTEC_Saisie:TEC_Ajoute_Ligne", startTime)

End Sub

Sub TEC_Modifie_Ligne() '2023-12-23 @ 07:04

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Modifie_Ligne", 0)

    If Fn_TEC_Is_Data_Valid() = False Then Exit Sub

    'Get the Client_ID
    ufSaisieHeures.txtClient_ID.value = Fn_GetID_From_Client_Name(ufSaisieHeures.txtClient.value)
    
    Call TEC_Record_Add_Or_Update_To_DB(ufSaisieHeures.txtTEC_ID.value)
    Call TEC_Record_Add_Or_Update_Locally(ufSaisieHeures.txtTEC_ID.value)
 
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

    Call TEC_Get_All_TEC_AF
    Call TEC_Refresh_ListBox_And_Add_Hours
    
    rmv_state = rmv_modeCreation
    
    ufSaisieHeures.txtClient.SetFocus
    
    Call Log_Record("modTEC_Saisie:TEC_Modifie_Ligne", startTime)

End Sub

Sub TEC_Efface_Ligne() '2023-12-23 @ 07:05

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Efface_Ligne", 0)

    If ufSaisieHeures.txtTEC_ID.value = "" Then
        MsgBox Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
            Buttons:=vbCritical
        GoTo Clean_Exit
    End If
    
    Dim answerYesNo As Long
    answerYesNo = MsgBox("Êtes-vous certain de vouloir DÉTRUIRE cet enregistrement ? ", _
                         vbYesNo + vbQuestion, "Confirmation de DESTRUCTION")
    If answerYesNo = vbNo Then
        MsgBox _
        Prompt:="Cet enregistrement ne sera PAS détruit ! ", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        GoTo Clean_Exit
    End If
    
    Call Log_Record("modTEC_Saisie:TEC_Efface_Ligne - Le DELETE est confirmé - " & CStr(-ufSaisieHeures.txtTEC_ID.value), -1) '2024-10-05 @ 07:21
    
    Dim Sh As Worksheet: Set Sh = wshTEC_Local
    
    Dim TECID As Long
    'With a negative ID value, it means to soft delete this record
    TECID = -ufSaisieHeures.txtTEC_ID.value
    
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
        Prompt:="L'enregistrement a été DÉTRUIT !", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        
    ufSaisieHeures.cmbProfessionnel.Enabled = True
    ufSaisieHeures.txtDate.Enabled = True
    rmv_state = rmv_modeCreation
    
    Call TEC_Get_All_TEC_AF
    
    Call TEC_Refresh_ListBox_And_Add_Hours
    
Clean_Exit:

    ufSaisieHeures.txtTEC_ID.value = ""
    ufSaisieHeures.txtClient.SetFocus

    'Libérer la mémoire
    Set Sh = Nothing

    Call Log_Record("modTEC_Saisie:TEC_Efface_Ligne", startTime)

End Sub

Sub TEC_Get_All_TEC_AF() '2024-11-19 @ 10:39
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Get_All_TEC_AF - " & _
                                        ufSaisieHeures.txtProf_ID.value & "/" & ufSaisieHeures.txtDate.value, 0)

    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    Application.ScreenUpdating = False

    'ProfID and Date are mandatory to execute this routine
    If ufSaisieHeures.txtProf_ID.value = "" Or ufSaisieHeures.txtDate.value = "" Then
        Exit Sub
    End If
    
    'Set criteria directly in TEC_Local for AdvancedFilter
    With ws
        .Range("R3").value = ufSaisieHeures.txtProf_ID.value
        .Range("S3").value = CLng(CDate(ufSaisieHeures.txtDate.value))
        .Range("T3").value = "FAUX"
    End With
    
    'Effacer les données de la dernière utilisation
    ws.Range("S6:S10").ClearContents
    ws.Range("S6").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_TEC_Local[#All]")
    ws.Range("S7").value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("R2:T3")
    ws.Range("S8").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("V1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("V2:AI2")
    ws.Range("S9").value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False

    Dim lastResultRow As Long
    lastResultRow = ws.Cells(ws.Rows.count, "V").End(xlUp).row
    ws.Range("S10").value = (lastResultRow - 2) & " lignes"
        
    If lastResultRow < 4 Then GoTo No_Sort_Required
    With ws.Sort 'Sort - Date / Prof / TEC_ID
        .SortFields.Clear
        .SortFields.Add key:=wshTEC_Local.Range("X3"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort Based On Date
        .SortFields.Add key:=wshTEC_Local.Range("W3"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort Based On Prof
        .SortFields.Add key:=wshTEC_Local.Range("V3"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort Based On Tec_ID
        .SetRange wshTEC_Local.Range("V3:AI" & lastResultRow) 'Set Range
        .Apply 'Apply Sort
     End With

No_Sort_Required:
    
    'Suddenly, I have to convert BOOLEAN value to TEXT !!!! - 2024-06-19 @ 14:20
    Dim r As Range
    If lastResultRow > 2 Then
        Set r = wshTEC_Local.Range("AC3:AC" & lastResultRow)
        Call ConvertRangeBooleanToText(r)
        Set r = wshTEC_Local.Range("AE3:AE" & lastResultRow)
        Call ConvertRangeBooleanToText(r)
        Set r = wshTEC_Local.Range("AG3:AG" & lastResultRow)
        Call ConvertRangeBooleanToText(r)
    End If
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set r = Nothing
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Get_All_TEC_AF", startTime)

End Sub

Sub TEC_Efface_Formulaire() 'Clear all fields on the userForm

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Efface_Formulaire", 0)

    'Empty the dynamic fields after reseting the form
    With ufSaisieHeures
        .txtTEC_ID.value = "" '2024-03-01 @ 09:56
        .txtClient.value = ""
        .txtClient_ID.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .txtSavedHeures = ""
        .cmbProfessionnel.Enabled = True
        .txtDate.Enabled = True
    End With
    
    Call TEC_Get_All_TEC_AF
    
    Call TEC_Refresh_ListBox_And_Add_Hours
    
    Call ActiverButtonsVraiOuFaux(False, False, False, False)
        
    ufSaisieHeures.txtClient.SetFocus
    
    Call Log_Record("modTEC_Saisie:TEC_Efface_Formulaire", startTime)

End Sub

Sub TEC_Record_Add_Or_Update_To_DB(TECID As Long) 'Write -OR- Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB(" & TECID & ")", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "TEC_Local$"
    
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
    
    Dim dateValue As Date '2024-09-04 @ 09:01
    dateValue = ufSaisieHeures.txtDate.value
    'Special log to debug Date Format issue... 2024-09-06 @ 16:32
    If TECID = 0 And Now - dateValue > 30 Then
        MsgBox "La date saisie est plus de 30 jours dans le passé..." & vbNewLine & vbNewLine & _
                "Veuillez aviser le développeur de cette situation SVP", vbInformation
        Call Log_Saisie_Heures("Future   ", "Plus de 30 jours dans le passé - dateValue = " & dateValue & "  type = " & TypeName(dateValue))
    End If
    
    If TECID < 0 Then 'Soft delete a record
        
        'Open the recordset for the specified ID
        
        rs.Open "SELECT * FROM [" & destinationTab & "] WHERE TEC_ID=" & Abs(TECID), conn, 2, 3
        saveLogTEC_ID = TECID
        If Not rs.EOF Then
            'Update the "IsDeleted" field to mark the record as deleted
            rs.Fields("DateSaisie").value = CDate(Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
            rs.Fields("EstDetruit").value = Fn_Convert_Value_Boolean_To_Text(True)
            rs.Fields("VersionApp").value = ThisWorkbook.Name
            rs.update
            
            Call Log_Saisie_Heures("DELETE" & saveLogTEC_ID, ufSaisieHeures.cmbProfessionnel.value & " | " & _
                                    dateValue & " | " & _
                                    ufSaisieHeures.txtClient_ID.value & " | " & _
                                    ufSaisieHeures.txtClient.value & " | " & _
                                    ufSaisieHeures.txtActivite.value & " | " & _
                                    Format$(ufSaisieHeures.txtHeures.value, "#0.00") & " | " & _
                                    Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value) & " | " & _
                                    ufSaisieHeures.txtCommNote.value)

        Else 'Handle the case where the specified ID is not found - PROBLEM !!!
            
            MsgBox "L'enregistrement avec le TEC_ID '" & TECID & "' ne peut être trouvé!", _
                vbExclamation
                
            rs.Close
            conn.Close
            
            Exit Sub
        End If
    
    Else 'Add a new record (TECID = 0) -OR- update an existing one (TECID <> 0)
        
        If TECID = 0 Then 'Add a record
        
            'SQL select command to find the next available ID
            Dim strSQL As String, MaxID As Long
            strSQL = "SELECT MAX(TEC_ID) AS MaxID FROM [" & destinationTab & "]"
        
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
            
            ufSaisieHeures.txtTEC_ID.value = nextID
            saveLogTEC_ID = nextID
        
            'Close the previous recordset, no longer needed and open an empty recordset
            rs.Close
            rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
            
            'Create a new RecordSet and update all fields of the recordset before updating it
            rs.AddNew
            rs.Fields("TEC_ID").value = nextID
            rs.Fields("Prof_ID").value = ufSaisieHeures.txtProf_ID.value
            rs.Fields("Prof").value = ufSaisieHeures.cmbProfessionnel.value
            rs.Fields("Date").value = dateValue '2024-09-04 @ 09:01
            rs.Fields("Client_ID").value = ufSaisieHeures.txtClient_ID.value
            rs.Fields("ClientNom").value = ufSaisieHeures.txtClient.value
            If Len(ufSaisieHeures.txtActivite.value) > 255 Then
                ufSaisieHeures.txtActivite.value = Left(ufSaisieHeures.txtActivite.value, 255)
            End If
            rs.Fields("Description").value = ufSaisieHeures.txtActivite.value
            rs.Fields("Heures").value = Format$(ufSaisieHeures.txtHeures.value, "#0.00")
            rs.Fields("CommentaireNote").value = ufSaisieHeures.txtCommNote.value
            rs.Fields("EstFacturable").value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value)
            rs.Fields("DateSaisie").value = CDate(Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
            rs.Fields("EstFacturee").value = Fn_Convert_Value_Boolean_To_Text(False)
            rs.Fields("DateFacturee").value = Null
            rs.Fields("EstDetruit").value = Fn_Convert_Value_Boolean_To_Text(False)
            rs.Fields("VersionApp").value = ThisWorkbook.Name
            rs.Fields("NoFacture").value = ""
            rs.update
            
            'Nouveau log - 2024-09-02 @ 10:40
            Call Log_Saisie_Heures("ADD    " & saveLogTEC_ID, ufSaisieHeures.cmbProfessionnel.value & " | " & _
                        dateValue & " | " & _
                        ufSaisieHeures.txtClient_ID.value & " | " & _
                        ufSaisieHeures.txtClient.value & " | " & _
                        ufSaisieHeures.txtActivite.value & " | " & _
                        Format$(ufSaisieHeures.txtHeures.value, "#0.00") & " | " & _
                        Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value) & " | " & _
                        ufSaisieHeures.txtCommNote.value)
        
        Else 'Update an existing record (TECID <> 0)
        
            'Open the recordset for the specified ID
            rs.Open "SELECT * FROM [" & destinationTab & "] WHERE TEC_ID=" & TECID, conn, 2, 3
            If Not rs.EOF Then
                'Update fields for the existing record
                rs.Fields("Client_ID").value = ufSaisieHeures.txtClient_ID.value
                rs.Fields("ClientNom").value = ufSaisieHeures.txtClient.value
                rs.Fields("Description").value = ufSaisieHeures.txtActivite.value
                rs.Fields("Heures").value = Format$(ufSaisieHeures.txtHeures.value, "#0.00")
                rs.Fields("CommentaireNote").value = ufSaisieHeures.txtCommNote.value
                rs.Fields("EstFacturable").value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value)
                rs.Fields("DateSaisie").value = CDate(Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
                rs.Fields("VersionApp").value = ThisWorkbook.Name
                'Nouveau log - 2024-09-02 @ 10:40
                
                Call Log_Saisie_Heures("UPDATE " & saveLogTEC_ID, ufSaisieHeures.cmbProfessionnel.value & " | " & _
                            dateValue & " | " & _
                            ufSaisieHeures.txtClient_ID.value & " | " & _
                            ufSaisieHeures.txtClient.value & " | " & _
                            ufSaisieHeures.txtActivite.value & " | " & _
                            Format$(ufSaisieHeures.txtHeures.value, "#0.00") & " | " & _
                            Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value) & " | " & _
                            ufSaisieHeures.txtCommNote.value)
            
            Else
            
                'Handle the case where the specified ID is not found - PROBLEM !!!
                
                MsgBox "L'enregistrement avec le TEC_ID '" & TECID & "' ne peut être trouvé!", vbExclamation
                Call Log_Record("#00480 - N'a pas trouvé le TECID '" & CStr(saveLogTEC_ID) & "'", -1) '2024-09-13 @ 09:09
                Call Log_Saisie_Heures("Erreur  ", "@00495 - Impossible de trouver le TEC_ID = " & CStr(saveLogTEC_ID)) '2024-09-02 @ 10:35
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
    conn.Close
    On Error GoTo 0
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB - " & TECID, startTime)

    Exit Sub
    
ErrorHandler:
     'Si une erreur survient, cela signifie que le fichier est en lecture seule
    MsgBox "Le fichier 'MASTER' est en lecture seule" & vbNewLine & vbNewLine & _
           "ou déjà ouvert par un autre utilisateur ou" & vbNewLine & vbNewLine & _
           "ou un autre type de problème" & vbNewLine & vbNewLine & _
           "COMMUNIQUER AVEC LE DÉVELOPPEUR IMMÉDIATEMENT", vbCritical, "Problème MAJEUR"
    If Not rs Is Nothing Then
        rs.Close
    End If
    If Not conn Is Nothing Then
        conn.Close
        On Error Resume Next
        Set conn = Nothing
        Set rs = Nothing
        On Error GoTo 0
    End If
    
    Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB - " & TECID, startTime)
    
End Sub

Sub TEC_Record_Add_Or_Update_Locally(TECID As Long) 'Write -OR- Update a record to local worksheet
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_Locally(" & TECID & ")", 0)

    Application.ScreenUpdating = False
    
    'What is the row number of this TEC_ID ?
    Dim lastUsedRow As Long
    
    Dim hoursValue As Double '2024-03-01 @ 05:40
    hoursValue = CDbl(ufSaisieHeures.txtHeures.value)
    
    Dim dateValue As Date
    dateValue = ufSaisieHeures.txtDate.value
    
    If TECID = 0 Then 'Add a new record
        'Get the next available row in TEC_Local
        Dim nextRowNumber As Long
        With wshTEC_Local
            nextRowNumber = .Cells(.Rows.count, 1).End(xlUp).row + 1
            .Range("A" & nextRowNumber).value = ufSaisieHeures.txtTEC_ID.value
            .Range("B" & nextRowNumber).value = ufSaisieHeures.txtProf_ID.value
            .Range("C" & nextRowNumber).value = ufSaisieHeures.cmbProfessionnel.value
            .Range("D" & nextRowNumber).value = dateValue
            .Range("E" & nextRowNumber).value = ufSaisieHeures.txtClient_ID.value
            .Range("F" & nextRowNumber).value = ufSaisieHeures.txtClient.value
            .Range("G" & nextRowNumber).value = ufSaisieHeures.txtActivite.value
            .Range("H" & nextRowNumber).value = hoursValue
            .Range("I" & nextRowNumber).value = ufSaisieHeures.txtCommNote.value
            .Range("J" & nextRowNumber).value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value)
            .Range("K" & nextRowNumber).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("L" & nextRowNumber).value = Fn_Convert_Value_Boolean_To_Text(False)
            .Range("M" & nextRowNumber).value = ""
            .Range("N" & nextRowNumber).value = Fn_Convert_Value_Boolean_To_Text(False)
            .Range("O" & nextRowNumber).value = ThisWorkbook.Name
            .Range("P" & nextRowNumber).value = ""
        End With
    Else
        'What is the row number for the TEC_ID
        lastUsedRow = wshTEC_Local.Cells(wshTEC_Local.Rows.count, "A").End(xlUp).row
        Dim lookupRange As Range: Set lookupRange = wshTEC_Local.Range("A3:A" & lastUsedRow)
        Dim rowToBeUpdated As Long
        rowToBeUpdated = Fn_Find_Row_Number_TEC_ID(Abs(TECID), lookupRange)
        If rowToBeUpdated < 1 Then
            'Handle the case where the specified TecID is not found !!
            MsgBox "L'enregistrement avec le TEC_ID '" & TECID & "' ne peut être trouvé!", _
                vbExclamation
            Exit Sub
        End If

        If TECID > 0 Then 'Modify the record
            With wshTEC_Local
                .Range("E" & rowToBeUpdated).value = ufSaisieHeures.txtClient_ID.value
                .Range("F" & rowToBeUpdated).value = ufSaisieHeures.txtClient.value
                .Range("G" & rowToBeUpdated).value = ufSaisieHeures.txtActivite.value
                .Range("H" & rowToBeUpdated).value = hoursValue
                .Range("I" & rowToBeUpdated).value = ufSaisieHeures.txtCommNote.value
                .Range("J" & rowToBeUpdated).value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value)
                .Range("K" & rowToBeUpdated).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
                .Range("L" & rowToBeUpdated).value = Fn_Convert_Value_Boolean_To_Text(False)
                .Range("M" & rowToBeUpdated).value = ""
                .Range("N" & rowToBeUpdated).value = Fn_Convert_Value_Boolean_To_Text(False)
                .Range("O" & rowToBeUpdated).value = ThisWorkbook.Name
                .Range("P" & rowToBeUpdated).value = ""
            End With
        Else 'Soft delete the record
            With wshTEC_Local
                .Range("K" & rowToBeUpdated).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
                .Range("N" & rowToBeUpdated).value = Fn_Convert_Value_Boolean_To_Text(True)
                .Range("O" & rowToBeUpdated).value = ThisWorkbook.Name
            End With
        End If
    End If
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_Locally - " & TECID, startTime)

End Sub

Sub TEC_Refresh_ListBox_And_Add_Hours() 'Load the listBox with the appropriate records

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Refresh_ListBox_And_Add_Hours - " & _
        ufSaisieHeures.txtProf_ID.value & "/" & ufSaisieHeures.txtDate.value, 0)

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
'    Application.Calculation = xlCalculationManual

   If ufSaisieHeures.txtProf_ID.value = "" Or Not IsDate(ufSaisieHeures.txtDate.value) Then
        MsgBox "Veuillez entrer un professionnel et/ou une date valide.", vbExclamation
        GoTo EndOfProcedure
    End If
    
'    'Modifie le critère pour forcer une execution du AdvancedFilter dans wshTEC_TDB_Data
'    wshTEC_TDB_Data.Range("S7").value = ufSaisieHeures.cmbProfessionnel.value
'
    ufSaisieHeures.txtTotalHeures.value = ""
    ufSaisieHeures.txtHresFact.value = ""
    ufSaisieHeures.txtHresNF.value = ""
    ufSaisieHeures.txtHresFactSemaine.value = ""
    ufSaisieHeures.txtHresNFSemaine.value = ""

    ufSaisieHeures.lsbHresJour.RowSource = ""
    ufSaisieHeures.lsbHresJour.Clear '2024-08-10 @ 05:59
    
    With ufSaisieHeures.lsbHresJour
        .ColumnHeads = False
        .ColumnCount = 9
        .ColumnWidths = "30; 24; 54; 157; 242; 35; 90; 32; 90"
    End With
    
    'Manually add to listBox (because some tests have to be made)
    Dim lastRow As Long
    lastRow = wshTEC_Local.Cells(wshTEC_Local.Rows.count, "V").End(xlUp).row
    Dim rng As Range
    Set rng = wshTEC_Local.Range("V3:AI" & lastRow)
     
    'Variables initiales
    Dim totalHeures As Currency: totalHeures = 0
    Dim totalHresFact As Currency: totalHresFact = 0
    Dim totalHresNonFact As Currency: totalHresNonFact = 0
    Dim rngResult As Range
    Dim i As Long, colIndex As Long
    
    'Remplissage du listBox
    If lastRow >= 3 Then
        Set rng = wshTEC_Local.Range("V3:AI" & lastRow)
        For i = 1 To rng.Rows.count
            With ufSaisieHeures.lsbHresJour
                .AddItem rng.Cells(i, 1).value
                For colIndex = 2 To 9
                    .List(.ListCount - 1, colIndex - 1) = rng.Cells(i, colIndex).value
                Next colIndex
            End With
            totalHeures = totalHeures + CCur(rng.Cells(i, 6).value)
            ' Calcul des heures facturables
            If Fn_Is_Client_Facturable(rng.Cells(i, 14).value) Then
                totalHresFact = totalHresFact + CCur(rng.Cells(i, 6).value)
            Else
                totalHresNonFact = totalHresNonFact + CCur(rng.Cells(i, 6).value)
            End If
        Next i
    End If

'    For i = 1 To rng.Rows.count
'        ufSaisieHeures.lsbHresJour.AddItem rng.Cells(i, 1).value
'        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 1) = rng.Cells(i, 2).value
'        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 2) = Format$(rng.Cells(i, 3).value, wshAdmin.Range("B1").value)
'        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 3) = rng.Cells(i, 4).value
'        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 4) = rng.Cells(i, 5).value
'        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 5) = Format$(rng.Cells(i, 6).value, "#,##0.00")
'        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 6) = rng.Cells(i, 7).value
'        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 7) = rng.Cells(i, 8).value
'        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 8) = Format$(rng.Cells(i, 9).value, wshAdmin.Range("B1").value & " hh:mm:ss")
'        totalHeures = totalHeures + CCur(rng.Cells(i, 6).value)
'        If Fn_Is_Client_Facturable(rng.Cells(i, 14).value) = True And rng.Cells(i, 8).value = "VRAI" Then
'            totalHresFact = totalHresFact + CCur(rng.Cells(i, 6).value)
'        Else
'            totalHresNonFact = totalHresNonFact + CCur(rng.Cells(i, 6).value)
'        End If
'    Next i
'    Application.ScreenUpdating = True

    'Mise à jour des totaux
    ufSaisieHeures.txtTotalHeures.value = Format$(totalHeures, "#0.00")
    ufSaisieHeures.txtHresFact.value = Format$(totalHresFact, "#0.00")
    ufSaisieHeures.txtHresNF.value = Format$(totalHresNonFact, "#0.00")
    
    'Maintenant, on traite la semaine à partir de wshTEC_TDB_Data
    Dim totalHresFactSemaine As Currency
    Dim totalHresNonFactSemaine As Currency
    
    'Modifie les critères pour forcer une execution du AdvancedFilter dans wshTEC_TDB_Data
    Dim dateCharge As Date, dateLundi As Date, dateDimanche As Date
    dateCharge = ufSaisieHeures.txtDate.value
    dateLundi = Fn_Obtenir_Date_Lundi(dateCharge)
    dateDimanche = dateLundi + 6
    Application.EnableEvents = False
    wshTEC_TDB_Data.Range("S7").value = ufSaisieHeures.cmbProfessionnel.value
    wshTEC_TDB_Data.Range("T7").value = dateLundi
    Application.EnableEvents = True
    wshTEC_TDB_Data.Range("U7").value = dateDimanche
    
    DoEvents
    
    lastRow = wshTEC_TDB_Data.Cells(wshTEC_TDB_Data.Rows.count, "W").End(xlUp).row
    If lastRow > 1 Then
        Set rngResult = wshTEC_TDB_Data.Range("W2:AD" & lastRow)
        totalHresFactSemaine = Application.WorksheetFunction.Sum(rngResult.Columns(7))
        totalHresNonFactSemaine = Application.WorksheetFunction.Sum(rngResult.Columns(8))
    End If

    ufSaisieHeures.txtHresFactSemaine.value = Format$(totalHresFactSemaine, "#0.00")
    ufSaisieHeures.txtHresNFSemaine.value = Format$(totalHresNonFactSemaine, "#0.00")
    
    ufSaisieHeures.Repaint
    
    DoEvents '2024-08-12 @ 10:31
    
    Application.ScreenUpdating = True

EndOfProcedure:

    Call ActiverButtonsVraiOuFaux(False, False, False, False)

    ufSaisieHeures.txtClient.SetFocus
    
    'Libération et fin
    Application.ScreenUpdating = True
    Application.EnableEvents = True
'    Application.Calculation = xlCalculationAutomatic
    Set rng = Nothing
    Set rngResult = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Refresh_ListBox_And_Add_Hours - " & _
                        ufSaisieHeures.txtProf_ID.value & "/" & ufSaisieHeures.txtDate.value, startTime)
    Exit Sub
    
ErrorHandler:

    MsgBox "Erreur : " & Err.Description, vbCritical
    Resume EndOfProcedure
    
End Sub

Sub TEC_Refresh_ListBox_And_Add_Hours_OK() 'Load the listBox with the appropriate records

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Refresh_ListBox_And_Add_Hours - " & _
        ufSaisieHeures.txtProf_ID.value & "/" & ufSaisieHeures.txtDate.value, 0)

    If ufSaisieHeures.txtProf_ID.value = "" Or ufSaisieHeures.txtDate.value = "" Then
        GoTo EndOfProcedure
    End If
    
'    'Modifie le critère pour forcer une execution du AdvancedFilter dans wshTEC_TDB_Data
'    wshTEC_TDB_Data.Range("S7").value = ufSaisieHeures.cmbProfessionnel.value
'
    ufSaisieHeures.txtTotalHeures.value = ""
    ufSaisieHeures.txtHresFact.value = ""
    ufSaisieHeures.txtHresNF.value = ""
    ufSaisieHeures.txtHresFactSemaine.value = ""
    ufSaisieHeures.txtHresNFSemaine.value = ""

    ufSaisieHeures.lsbHresJour.RowSource = ""
    ufSaisieHeures.lsbHresJour.Clear '2024-08-10 @ 05:59
    
    'Last Row used in first column of result
    Dim lastRow As Long
    lastRow = wshTEC_Local.Cells(wshTEC_Local.Rows.count, "V").End(xlUp).row
    If lastRow < 3 Then GoTo Rien_Aujourdhui
        
    With ufSaisieHeures.lsbHresJour
        .ColumnHeads = False
        .ColumnCount = 9
        .ColumnWidths = "30; 24; 54; 157; 242; 35; 90; 32; 90"
    End With
    
    'Manually add to listBox (because some tests have to be made)
    Dim rng As Range
    Set rng = wshTEC_Local.Range("V3:AI" & lastRow)
     
    Dim i As Long, j As Long
    Dim totalHeures As Currency
    Dim totalHresFact As Currency, totalHresNonFact As Currency
    Application.ScreenUpdating = False
    For i = 1 To rng.Rows.count
        ufSaisieHeures.lsbHresJour.AddItem rng.Cells(i, 1).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 1) = rng.Cells(i, 2).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 2) = Format$(rng.Cells(i, 3).value, wshAdmin.Range("B1").value)
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 3) = rng.Cells(i, 4).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 4) = rng.Cells(i, 5).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 5) = Format$(rng.Cells(i, 6).value, "#,##0.00")
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 6) = rng.Cells(i, 7).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 7) = rng.Cells(i, 8).value
        ufSaisieHeures.lsbHresJour.List(ufSaisieHeures.lsbHresJour.ListCount - 1, 8) = Format$(rng.Cells(i, 9).value, wshAdmin.Range("B1").value & " hh:mm:ss")
        totalHeures = totalHeures + CCur(rng.Cells(i, 6).value)
        If Fn_Is_Client_Facturable(rng.Cells(i, 14).value) = True And rng.Cells(i, 8).value = "VRAI" Then
            totalHresFact = totalHresFact + CCur(rng.Cells(i, 6).value)
        Else
            totalHresNonFact = totalHresNonFact + CCur(rng.Cells(i, 6).value)
        End If
    Next i
    Application.ScreenUpdating = True
    
Rien_Aujourdhui:

    ufSaisieHeures.txtTotalHeures.value = Format$(totalHeures, "#0.00")
    ufSaisieHeures.txtHresFact.value = Format$(totalHresFact, "#0.00")
    ufSaisieHeures.txtHresNF.value = Format$(totalHresNonFact, "#0.00")
    
    'Maintenant, on traite la semaine à partir de wshTEC_TDB_Data
    Dim totalHresFactSemaine As Currency
    Dim totalHresNonFactSemaine As Currency
    
    'Modifie les critères pour forcer une execution du AdvancedFilter dans wshTEC_TDB_Data
    Dim dateCharge As Date, dateLundi As Date, dateDimanche As Date
    dateCharge = ufSaisieHeures.txtDate.value
    dateLundi = Fn_Obtenir_Date_Lundi(dateCharge)
    dateDimanche = dateLundi + 6
    Application.EnableEvents = False
    wshTEC_TDB_Data.Range("S7").value = ufSaisieHeures.cmbProfessionnel.value
    wshTEC_TDB_Data.Range("T7").value = dateLundi
    Application.EnableEvents = True
    wshTEC_TDB_Data.Range("U7").value = dateDimanche
    
    DoEvents
    
    lastRow = wshTEC_TDB_Data.Cells(wshTEC_TDB_Data.Rows.count, "W").End(xlUp).row
    If lastRow > 1 Then
        Dim rngResult As Range
        Set rngResult = wshTEC_TDB_Data.Range("W2:AD" & lastRow)
        totalHresFactSemaine = Application.WorksheetFunction.Sum(rngResult.Columns(7))
        totalHresNonFactSemaine = Application.WorksheetFunction.Sum(rngResult.Columns(8))
    End If

    ufSaisieHeures.txtHresFactSemaine.value = Format$(totalHresFactSemaine, "#0.00")
    ufSaisieHeures.txtHresNFSemaine.value = Format$(totalHresNonFactSemaine, "#0.00")
    
    ufSaisieHeures.Repaint
    
    DoEvents '2024-08-12 @ 10:31
    
    Application.ScreenUpdating = True

EndOfProcedure:

    Call ActiverButtonsVraiOuFaux(False, False, False, False)

    ufSaisieHeures.txtClient.SetFocus
    
    'Libérer la mémoire
    Set rng = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Refresh_ListBox_And_Add_Hours - " & _
                        ufSaisieHeures.txtProf_ID.value & "/" & ufSaisieHeures.txtDate.value, startTime)
    
End Sub

Sub TEC_Update_TDB_From_TEC_Local()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Update_TDB_From_TEC_Local", 0)

    Dim wsFrom As Worksheet: Set wsFrom = wshTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = wsFrom.Cells(wsFrom.Rows.count, 1).End(xlUp).row
    
    'Charger en mémoire toutes les données source
    Dim rawData As Variant
    rawData = wsFrom.Range("A3:N" & lastUsedRow).value
    
    'Préparer le tableau des données à la sortie
    Dim arr() As Variant
    Dim numRows As Long: numRows = UBound(rawData, 1)
    ReDim arr(1 To numRows, 1 To 11)
    
    Dim i As Long
    For i = 1 To numRows
        arr(i, 1) = rawData(i, 1) 'TEC_ID
        arr(i, 2) = Format$(rawData(i, 2), "000") 'ProfID
        arr(i, 3) = rawData(i, 3) 'Prof
        arr(i, 4) = rawData(i, 4) 'Date
        arr(i, 5) = rawData(i, 5) 'Client's ID
        arr(i, 6) = rawData(i, 6) 'Client's Name
        arr(i, 7) = IIf(Fn_Is_Client_Facturable(rawData(i, 5)), "VRAI", "FAUX") 'Facturable
        arr(i, 8) = rawData(i, 8) 'Hours
        arr(i, 9) = rawData(i, 10) 'isBillable
        arr(i, 10) = rawData(i, 12) 'isInvoiced
        arr(i, 11) = rawData(i, 14) 'isDeleted
    Next i
    
    ' Mettre à jour la feuille TEC_TDB_Data
    Dim rngTo As Range
    Set rngTo = wshTEC_TDB_Data.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2))
    rngTo.value = arr
    
    'Libérer la mémoire
    Set rngTo = Nothing
    Set wsFrom = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Update_TDB_From_TEC_Local", startTime)

End Sub

Sub TEC_Update_TDB_From_TEC_Local_OK()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Update_TDB_From_TEC_Local", 0)

    Dim wsFrom As Worksheet: Set wsFrom = wshTEC_Local
    
    Dim lastUsedRow As Long
    lastUsedRow = wsFrom.Cells(wsFrom.Rows.count, 1).End(xlUp).row
    
    Dim arr() As Variant
    ReDim arr(1 To lastUsedRow - 2, 1 To 11) '2 rows of Heading
    
    Dim i As Long
    Dim b As Boolean
    For i = 3 To lastUsedRow
        With wsFrom
            arr(i - 2, 1) = .Range("A" & i).value 'TEC_ID
            arr(i - 2, 2) = Format$(.Range("B" & i).value, "000") 'ProfID
            arr(i - 2, 3) = .Range("C" & i).value 'Prof
            arr(i - 2, 4) = .Range("D" & i).value 'Date
            arr(i - 2, 5) = .Range("E" & i).value 'Client's ID - 2024-09-24 @ 09:41
            arr(i - 2, 6) = .Range("F" & i).value 'Client's Name
            b = Fn_Is_Client_Facturable(.Range("E" & i).value)
            If UCase(b) = True Then
                 arr(i - 2, 7) = "VRAI"
            Else
                 arr(i - 2, 7) = "FAUX"
            End If
            arr(i - 2, 8) = .Range("H" & i).value 'Hours
            arr(i - 2, 9) = .Range("J" & i).value 'isBillable
            arr(i - 2, 10) = .Range("L" & i).value 'isInvoiced
            arr(i - 2, 11) = .Range("N" & i).value 'isDeleted
        End With
    Next i

    'Mettre à jour la feuille TEC_TDB_Data au complet
    Dim rngTo As Range
    Set rngTo = wshTEC_TDB_Data.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2))
    rngTo.value = arr
    
    'Libérer la mémoire
    Set rngTo = Nothing
    Set wsFrom = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Update_TDB_From_TEC_Local", startTime)

End Sub

Sub TEC_TdB_Refresh_All_Pivot_Tables()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_TdB_Refresh_All_Pivot_Tables", 0)

    Dim pt As pivotTable
    For Each pt In wshTEC_TDB_PivotTable.PivotTables
        pt.RefreshTable
    Next pt

    'Libérer la mémoire
    Set pt = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_TdB_Refresh_All_Pivot_Tables", startTime)
    
End Sub

Sub ActiverButtonsVraiOuFaux(a As Boolean, u As Boolean, _
                                  d As Boolean, c As Boolean)
                                  
    With ufSaisieHeures
        .cmdAdd.Enabled = a
        .cmdUpdate.Enabled = u
        .cmdDelete.Enabled = d
        .cmdClear.Enabled = c
    End With

End Sub

Sub AfficherMessageDateInvalide(location As String) '2024-06-13 @ 12:40

    MsgBox "La date saisie ne peut être acceptée tel qu'elle est entrée." & vbNewLine & vbNewLine & _
           "Elle doit être obligatoirement de format:" & vbNewLine & _
           "     'j', jj', " & vbNewLine & _
           "     'jj-mm', 'jj/mm' ou " & vbNewLine & _
           "     'j-m-aa', 'j-m-aaaa', 'jj-mm-aaaa'" & vbNewLine & vbNewLine & _
           "Veuillez saisir la date de nouveau SVP", _
           vbCritical, _
           "La date saisie est INVALIDE - " & location

End Sub

Sub UpdatePivotTables()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:UpdatePivotTables", 0)
    
    Dim ws As Worksheet: Set ws = wshStatsHeuresPivotTables
    Dim pt As pivotTable
    
    'Parcourt tous les PivotTables dans chaque feuille
    For Each pt In ws.PivotTables
        On Error Resume Next
        Application.EnableEvents = False
        pt.PivotCache.Refresh 'Actualise le cache Pivot
        Application.EnableEvents = True
        On Error GoTo 0
    Next pt

    'Libérer la mémoire
    Set pt = Nothing
    Set ws = Nothing
    
    Call Log_Record("modAppli:UpdatePivotTables", startTime)
    
End Sub


