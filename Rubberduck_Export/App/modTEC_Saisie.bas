Attribute VB_Name = "modTEC_Saisie"
Option Explicit

Public Const rmv_modeInitial As Long = 1
Public Const rmv_modeCreation As Long = 2
Public Const rmv_modeAffichage As Long = 3
Public Const rmv_modeModification As Long = 4

Public rmv_state As Long

Public savedClient As String
Public savedActivite As String
Public savedHeures As Currency
Public savedFacturable As String
Public savedCommNote As String

Sub TEC_Ajoute_Ligne() 'Add an entry to DB

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Ajoute_Ligne", "", 0)

    'Obtenir le ID du client pur (à partir de son nom pur)
    ufSaisieHeures.txtClientID.value = Fn_Cell_From_BD_Client(ufSaisieHeures.txtClient.value, 1, 2)
        
    If Fn_TEC_Is_Data_Valid() = True Then
        Dim y As Integer, m As Integer, d As Integer
        Dim avant As String
        On Error Resume Next
            avant = ufSaisieHeures.txtDate.value
            y = year(ufSaisieHeures.txtDate.value)
            m = month(ufSaisieHeures.txtDate.value)
            d = day(ufSaisieHeures.txtDate.value)
'            If y = 2024 And m < 9 Then 'Si mois < 9 alors, on prend pour acquis que le jour et le mois sont inversés...
'                Dim temp As Integer
'                temp = m
'                m = d
'                d = temp
'                Call Log_Saisie_Heures("info     ", "@00045 - AJUSTEMENT (PLUG) --->   y = " & y & "   m = " & m & "   d = " & d & "   type = " & TypeName(ufSaisieHeures.txtDate.value))
'            End If
            ufSaisieHeures.txtDate.value = Format$(DateSerial(y, m, d), "yyyy-mm-dd")
        On Error GoTo 0
        
        Call TEC_Record_Add_Or_Update_To_DB(0)
        
        Call TEC_Record_Add_Or_Update_Locally(0)
        
        'Clear the userForm fields after saving
        With ufSaisieHeures
            .txtTECID.value = ""
            .txtClient.value = ""
            .txtClientID.value = ""
            .txtActivite.value = ""
            .txtHeures.value = ""
            .txtCommNote.value = ""
            .chbFacturable = True
'            .txtSavedHeures.value = ""
        End With
        
        Call TEC_Get_All_TEC_AF
        
        Call TEC_Refresh_ListBox_And_Add_Hours
        
        'Reset command buttons
        Call ActiverButtonsVraiOuFaux(False, False, False, False)
        
        'Back to client
        ufSaisieHeures.txtClient.SetFocus
    End If
    
    Call Log_Record("modTEC_Saisie:TEC_Ajoute_Ligne", "", startTime)

End Sub

Sub TEC_Modifie_Ligne() '2023-12-23 @ 07:04

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Modifie_Ligne", "", 0)

    If Fn_TEC_Is_Data_Valid() = False Then Exit Sub

    'Obtenir le ID du client pur (à partir de son nom pur) - 2025-03-04 @ 08:02
    ufSaisieHeures.txtClientID.value = Fn_Cell_From_BD_Client(ufSaisieHeures.txtClient.value, 1, 2)
        
    Call TEC_Record_Add_Or_Update_To_DB(ufSaisieHeures.txtTECID.value)
    Call TEC_Record_Add_Or_Update_Locally(ufSaisieHeures.txtTECID.value)
 
    'Initialize dynamic variables
    With ufSaisieHeures
        .txtTECID.value = ""
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
    
    Call Log_Record("modTEC_Saisie:TEC_Modifie_Ligne", "", startTime)

End Sub

Sub TEC_Efface_Ligne() '2023-12-23 @ 07:05

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Efface_Ligne", "", 0)

    If ufSaisieHeures.txtTECID.value = "" Then
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
    
    Call Log_Record("modTEC_Saisie:TEC_Efface_Ligne - Le DELETE est confirmé - " & CStr(-ufSaisieHeures.txtTECID.value), -1) '2024-10-05 @ 07:21
    
    Dim Sh As Worksheet: Set Sh = wsdTEC_Local
    
    Dim tecID As Long
    'With a negative ID value, it means to soft delete this record
    tecID = -ufSaisieHeures.txtTECID.value
    
    Call TEC_Record_Add_Or_Update_To_DB(tecID)  'Write to external XLSX file - 2023-12-23 @ 07:07
    Call TEC_Record_Add_Or_Update_Locally(tecID)  'Write to local worksheet - 2024-02-25 @ 10:40
    
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

    ufSaisieHeures.txtTECID.value = ""
    ufSaisieHeures.txtClient.SetFocus

    'Libérer la mémoire
    Set Sh = Nothing

    Call Log_Record("modTEC_Saisie:TEC_Efface_Ligne", "", startTime)

End Sub

Sub TEC_Get_All_TEC_AF() '2024-11-19 @ 10:39
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Get_All_TEC_AF", _
                                                                 ufSaisieHeures.txtProfID.value & "/" & ufSaisieHeures.txtDate.value, 0)

    Dim ws As Worksheet: Set ws = wsdTEC_Local
    
    Application.ScreenUpdating = False

    'ProfID and Date are mandatory to execute this routine
    If ufSaisieHeures.txtProfID.value = "" Or ufSaisieHeures.txtDate.value = "" Then
        Exit Sub
    End If
    
    'wsdTEC_Local_AF#1

    'Set criteria directly in TEC_Local for AdvancedFilter
    With ws
        .Range("R3").value = ufSaisieHeures.txtProfID.value
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
    With ws.Sort 'Sort - Date / Prof / TECID
        .SortFields.Clear
        .SortFields.Add key:=wsdTEC_Local.Range("X3"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort Based On Date
        .SortFields.Add key:=wsdTEC_Local.Range("W3"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort Based On Prof
        .SortFields.Add key:=wsdTEC_Local.Range("V3"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort Based On TECID
        .SetRange wsdTEC_Local.Range("V3:AI" & lastResultRow) 'Set Range
        .Apply 'Apply Sort
     End With

No_Sort_Required:
    
    'Suddenly, I have to convert BOOLEAN value to TEXT !!!! - 2024-06-19 @ 14:20
    Dim r As Range
    If lastResultRow > 2 Then
        Set r = wsdTEC_Local.Range("AC3:AC" & lastResultRow)
        Call ConvertRangeBooleanToText(r)
        Set r = wsdTEC_Local.Range("AE3:AE" & lastResultRow)
        Call ConvertRangeBooleanToText(r)
        Set r = wsdTEC_Local.Range("AG3:AG" & lastResultRow)
        Call ConvertRangeBooleanToText(r)
    End If
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set r = Nothing
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Get_All_TEC_AF", "", startTime)

End Sub

Sub TEC_Efface_Formulaire() 'Clear all fields on the userForm

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Efface_Formulaire", "", 0)

    'Empty the dynamic fields after reseting the form
    With ufSaisieHeures
        .txtTECID.value = "" '2024-03-01 @ 09:56
        .txtClient.value = ""
        .txtClientID.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
'        .txtSavedHeures = ""
        .cmbProfessionnel.Enabled = True
        .txtDate.Enabled = True
    End With
    
    savedHeures = 0
    
    Call TEC_Get_All_TEC_AF
    
    Call TEC_Refresh_ListBox_And_Add_Hours
    
    Call ActiverButtonsVraiOuFaux(False, False, False, False)
        
    ufSaisieHeures.txtClient.SetFocus
    
    Call Log_Record("modTEC_Saisie:TEC_Efface_Formulaire", "", startTime)

End Sub

Sub TEC_Record_Add_Or_Update_To_DB(tecID As Long) 'Write -OR- Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB", CStr(tecID), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "TEC_Local$"
    
'    On Error GoTo ErrorHandler
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    Dim strConnection As String
    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
        destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    conn.Open strConnection
    
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim saveLogTECID As Long
    saveLogTECID = tecID
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim dateValue As Date '2024-09-04 @ 09:01
    dateValue = ufSaisieHeures.txtDate.value
    'Special log to debug Date Format issue... 2024-09-06 @ 16:32
    If tecID = 0 And Date - dateValue > 30 Then
        MsgBox "La date saisie est plus de 30 jours dans le passé..." & vbNewLine & vbNewLine & _
                "Veuillez aviser le développeur de cette situation SVP", vbInformation
        Call Log_Saisie_Heures("Future   ", "Plus de 30 jours dans le passé - dateValue = " & dateValue & "  type = " & TypeName(dateValue))
    End If
    
    If tecID < 0 Then 'Soft delete a record
        
        'Open the recordset for the specified ID
        
        rs.Open "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & Abs(tecID), conn, 2, 3
        saveLogTECID = tecID
        If Not rs.EOF Then
            'Update the "IsDeleted" field to mark the record as deleted
            rs.Fields(fTECDateSaisie - 1).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rs.Fields(fTECEstDetruit - 1).value = Fn_Convert_Value_Boolean_To_Text(True)
            rs.Fields(fTECVersionApp - 1).value = ThisWorkbook.Name
            rs.Update
            
            Call Log_Saisie_Heures("DELETE" & saveLogTECID, ufSaisieHeures.cmbProfessionnel.value & " | " & _
                                    dateValue & " | " & _
                                    ufSaisieHeures.txtClientID.value & " | " & _
                                    ufSaisieHeures.txtClient.value & " | " & _
                                    ufSaisieHeures.txtActivite.value & " | " & _
                                    Format$(ufSaisieHeures.txtHeures.value, "#0.00") & " | " & _
                                    Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value) & " | " & _
                                    ufSaisieHeures.txtCommNote.value)

        Else 'Handle the case where the specified ID is not found - PROBLEM !!!
            
            MsgBox "L'enregistrement avec le TECID '" & tecID & "' ne peut être trouvé!", _
                vbExclamation
                
            rs.Close
            conn.Close
            
            Exit Sub
        End If
    
    Else 'Add a new record (TECID = 0) -OR- update an existing one (TECID <> 0)
        
        If tecID = 0 Then 'Add a record
        
            'SQL select command to find the next available ID
            Dim strSQL As String, MaxID As Long
            strSQL = "SELECT MAX(TECID) AS MaxID FROM [" & destinationTab & "]"
        
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
            
            ufSaisieHeures.txtTECID.value = nextID
            saveLogTECID = nextID
        
            'Close the previous recordset, no longer needed and open an empty recordset
            rs.Close
            rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
            
            'Create a new RecordSet and update all fields of the recordset before updating it
            rs.AddNew
            rs.Fields(fTECTECID - 1).value = nextID
            rs.Fields(fTECProfID - 1).value = ufSaisieHeures.txtProfID.value
            rs.Fields(fTECProf - 1).value = ufSaisieHeures.cmbProfessionnel.value
            rs.Fields(fTECDate - 1).value = dateValue '2024-09-04 @ 09:01
            rs.Fields(fTECClientID - 1).value = ufSaisieHeures.txtClientID.value
            rs.Fields(fTECClientNom - 1).value = ufSaisieHeures.txtClient.value
            If Len(ufSaisieHeures.txtActivite.value) > 255 Then
                ufSaisieHeures.txtActivite.value = Left$(ufSaisieHeures.txtActivite.value, 255)
            End If
            rs.Fields(fTECDescription - 1).value = ufSaisieHeures.txtActivite.value
            rs.Fields(fTECHeures - 1).value = Format$(ufSaisieHeures.txtHeures.value, "#0.00")
            rs.Fields(fTECCommentaireNote - 1).value = ufSaisieHeures.txtCommNote.value
            rs.Fields(fTECEstFacturable - 1).value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value)
            rs.Fields(fTECDateSaisie - 1).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rs.Fields(fTECEstFacturee - 1).value = Fn_Convert_Value_Boolean_To_Text(False)
            rs.Fields(fTECDateFacturee - 1).value = Null
            rs.Fields(fTECEstDetruit - 1).value = Fn_Convert_Value_Boolean_To_Text(False)
            rs.Fields(fTECVersionApp - 1).value = ThisWorkbook.Name
            rs.Fields(fTECNoFacture - 1).value = ""
            rs.Update
            
            'Nouveau log - 2024-09-02 @ 10:40
            Call Log_Saisie_Heures("ADD    " & saveLogTECID, ufSaisieHeures.cmbProfessionnel.value & " | " & _
                        dateValue & " | " & _
                        ufSaisieHeures.txtClientID.value & " | " & _
                        ufSaisieHeures.txtClient.value & " | " & _
                        ufSaisieHeures.txtActivite.value & " | " & _
                        Format$(ufSaisieHeures.txtHeures.value, "#0.00") & " | " & _
                        Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value) & " | " & _
                        ufSaisieHeures.txtCommNote.value)
        
        Else 'Update an existing record (TECID <> 0)
        
            'Open the recordset for the specified ID
            rs.Open "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & tecID, conn, 2, 3
            If Not rs.EOF Then
                'Update fields for the existing record
                rs.Fields(fTECClientID - 1).value = ufSaisieHeures.txtClientID.value
                rs.Fields(fTECClientNom - 1).value = ufSaisieHeures.txtClient.value
                rs.Fields(fTECDescription - 1).value = ufSaisieHeures.txtActivite.value
                rs.Fields(fTECHeures - 1).value = Format$(ufSaisieHeures.txtHeures.value, "#0.00")
                rs.Fields(fTECCommentaireNote - 1).value = ufSaisieHeures.txtCommNote.value
                rs.Fields(fTECEstFacturable - 1).value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value)
                rs.Fields(fTECDateSaisie - 1).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
                rs.Fields(fTECVersionApp - 1).value = ThisWorkbook.Name
                
                Call Log_Saisie_Heures("UPDATE " & saveLogTECID, ufSaisieHeures.cmbProfessionnel.value & " | " & _
                            dateValue & " | " & _
                            ufSaisieHeures.txtClientID.value & " | " & _
                            ufSaisieHeures.txtClient.value & " | " & _
                            ufSaisieHeures.txtActivite.value & " | " & _
                            Format$(ufSaisieHeures.txtHeures.value, "#0.00") & " | " & _
                            Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value) & " | " & _
                            ufSaisieHeures.txtCommNote.value)
            
            Else
            
                'Handle the case where the specified ID is not found - PROBLEM !!!
                
                MsgBox "L'enregistrement avec le TECID '" & tecID & "' ne peut être trouvé!", vbExclamation
                Call Log_Record("ERREUR - N'a pas trouvé le TECID '", CStr(saveLogTECID), -1)   '2024-09-13 @ 09:09
                Call Log_Saisie_Heures("Erreur  ", "@00495 - Impossible de trouver le TECID = " & CStr(saveLogTECID)) '2024-09-02 @ 10:35
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
    conn.Close
    On Error GoTo 0
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB", CStr(tecID), startTime)

    Exit Sub
    
ErrorHandler:
     'Si une erreur survient, cela signifie que le fichier est en lecture seule
    MsgBox "Le fichier 'MASTER' est en lecture seule" & vbNewLine & vbNewLine & _
           "ou déjà ouvert par un autre utilisateur ou" & vbNewLine & vbNewLine & _
           "ou un autre type de problème" & vbNewLine & vbNewLine & _
           "COMMUNIQUER AVEC LE DÉVELOPPEUR IMMÉDIATEMENT", vbCritical, "Erreur = " & Err & " - " & Err.Description
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
    
    Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB", CStr(tecID), startTime)
    
End Sub

Sub TEC_Record_Add_Or_Update_Locally(tecID As Long) 'Write -OR- Update a record to local worksheet
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_Locally", CStr(tecID), 0)

    Application.ScreenUpdating = False
    
    'What is the row number of this TECID ?
    Dim lastUsedRow As Long
    
    Dim hoursValue As Double '2024-03-01 @ 05:40
    hoursValue = CDbl(ufSaisieHeures.txtHeures.value)
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim dateValue As Date
    dateValue = ufSaisieHeures.txtDate.value
    
    If tecID = 0 Then 'Add a new record
        'Get the next available row in TEC_Local
        Dim nextRowNumber As Long
        With wsdTEC_Local
            nextRowNumber = .Cells(.Rows.count, 1).End(xlUp).row + 1
            .Range("A" & nextRowNumber).value = ufSaisieHeures.txtTECID.value
            .Range("B" & nextRowNumber).value = ufSaisieHeures.txtProfID.value
            .Range("C" & nextRowNumber).value = ufSaisieHeures.cmbProfessionnel.value
            .Range("D" & nextRowNumber).value = dateValue
            .Range("E" & nextRowNumber).value = ufSaisieHeures.txtClientID.value
            .Range("F" & nextRowNumber).value = ufSaisieHeures.txtClient.value
            .Range("G" & nextRowNumber).value = ufSaisieHeures.txtActivite.value
            .Range("H" & nextRowNumber).value = hoursValue
            .Range("I" & nextRowNumber).value = ufSaisieHeures.txtCommNote.value
            .Range("J" & nextRowNumber).value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value)
            .Range("K" & nextRowNumber).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            .Range("L" & nextRowNumber).value = Fn_Convert_Value_Boolean_To_Text(False)
            .Range("M" & nextRowNumber).value = ""
            .Range("N" & nextRowNumber).value = Fn_Convert_Value_Boolean_To_Text(False)
            .Range("O" & nextRowNumber).value = ThisWorkbook.Name
            .Range("P" & nextRowNumber).value = ""
        End With
    Else
        'What is the row number for the TECID
        lastUsedRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "A").End(xlUp).row
        Dim lookupRange As Range: Set lookupRange = wsdTEC_Local.Range("A3:A" & lastUsedRow)
        Dim rowToBeUpdated As Long
        rowToBeUpdated = Fn_Find_Row_Number_TECID(Abs(tecID), lookupRange)
        If rowToBeUpdated < 1 Then
            'Handle the case where the specified TecID is not found !!
            MsgBox "L'enregistrement avec le TECID '" & tecID & "' ne peut être trouvé!", _
                vbExclamation
            Exit Sub
        End If

        If tecID > 0 Then 'Modify the record
            With wsdTEC_Local
                .Range("E" & rowToBeUpdated).value = ufSaisieHeures.txtClientID.value
                .Range("F" & rowToBeUpdated).value = ufSaisieHeures.txtClient.value
                .Range("G" & rowToBeUpdated).value = ufSaisieHeures.txtActivite.value
                .Range("H" & rowToBeUpdated).value = hoursValue
                .Range("I" & rowToBeUpdated).value = ufSaisieHeures.txtCommNote.value
                .Range("J" & rowToBeUpdated).value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chbFacturable.value)
                .Range("K" & rowToBeUpdated).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
                .Range("L" & rowToBeUpdated).value = Fn_Convert_Value_Boolean_To_Text(False)
                .Range("M" & rowToBeUpdated).value = ""
                .Range("N" & rowToBeUpdated).value = Fn_Convert_Value_Boolean_To_Text(False)
                .Range("O" & rowToBeUpdated).value = ThisWorkbook.Name
                .Range("P" & rowToBeUpdated).value = ""
            End With
        Else 'Soft delete the record
            With wsdTEC_Local
                .Range("K" & rowToBeUpdated).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
                .Range("N" & rowToBeUpdated).value = Fn_Convert_Value_Boolean_To_Text(True)
                .Range("O" & rowToBeUpdated).value = ThisWorkbook.Name
            End With
        End If
    End If
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_Locally", CStr(tecID), startTime)

End Sub

Sub TEC_Refresh_ListBox_And_Add_Hours() 'Load the listBox with the appropriate records

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Refresh_ListBox_And_Add_Hours", _
            ufSaisieHeures.txtProfID.value & "/" & ufSaisieHeures.txtDate.value, 0)

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If ufSaisieHeures.txtProfID.value = "" Or Not IsDate(ufSaisieHeures.txtDate.value) Then
        MsgBox "Veuillez entrer un professionnel et/ou une date valide.", vbExclamation
        GoTo EndOfProcedure
    End If
    
    'On vide le formulaire
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
        .ColumnWidths = "30; 23; 60; 157; 242; 35; 90; 32; 90"
    End With
    
    'Manually add to listBox (because some tests have to be made)
    Dim lastRow As Long
    lastRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "V").End(xlUp).row
    Dim rng As Range
    Set rng = wsdTEC_Local.Range("V3:AI" & lastRow)
     
    'Variables initiales
    Dim totalHeures As Currency: totalHeures = 0
    Dim totalHresFact As Currency: totalHresFact = 0
    Dim totalHresNonFact As Currency: totalHresNonFact = 0
    Dim rngResult As Range
    Dim i As Long, ColIndex As Long
    
    'Remplissage du listBox
    Dim hresFormat As String
    If lastRow >= 3 Then
        Set rng = wsdTEC_Local.Range("V3:AI" & lastRow)
        For i = 1 To rng.Rows.count
            With ufSaisieHeures.lsbHresJour
                .AddItem rng.Cells(i, 1).value
                For ColIndex = 2 To 9
                    If ColIndex <> 6 Then '2025-01-31 @ 14:42
                        .List(.ListCount - 1, ColIndex - 1) = rng.Cells(i, ColIndex).value
                    Else
                        hresFormat = Format$(rng.Cells(i, ColIndex).value, "#0.00")
                        hresFormat = Space(5 - Len(hresFormat)) & hresFormat
                        .List(.ListCount - 1, ColIndex - 1) = hresFormat
                    End If
                Next ColIndex
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
    
    Call Log_Record("modTEC_Saisie:TEC_Refresh_ListBox_And_Add_Hours", _
                        ufSaisieHeures.txtProfID.value & "/" & ufSaisieHeures.txtDate.value, startTime)
    Exit Sub
    
ErrorHandler:

    MsgBox "Erreur : " & Err.Description, vbCritical, "Erreur # APP-003"
    Resume EndOfProcedure
    
End Sub

Sub TEC_Update_TDB_From_TEC_Local()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Update_TDB_From_TEC_Local", "", 0)

    Dim wsFrom As Worksheet: Set wsFrom = wsdTEC_Local
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
        arr(i, 1) = rawData(i, 1) 'TECID
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
    
    Call Log_Record("modTEC_Saisie:TEC_Update_TDB_From_TEC_Local", "", startTime)

End Sub

Sub TEC_TdB_Refresh_All_Pivot_Tables()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_TdB_Refresh_All_Pivot_Tables", "", 0)

    Dim pt As pivotTable
    For Each pt In wshTEC_TDB.PivotTables
        pt.RefreshTable
    Next pt

    'Libérer la mémoire
    Set pt = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_TdB_Refresh_All_Pivot_Tables", "", startTime)
    
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

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:UpdatePivotTables", "", 0)
    
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
    
    Call Log_Record("modTEC_Saisie:UpdatePivotTables", "", startTime)
    
End Sub


