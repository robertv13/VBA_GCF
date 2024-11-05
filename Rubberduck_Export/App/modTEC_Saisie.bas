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

    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_Ajoute_Ligne @00018", True)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Ajoute_Ligne", 0)

    If Fn_TEC_Is_Data_Valid() = True Then
'        'Get the Client_ID
'        ufSaisieHeures.txtClient_ID.value = Fn_GetID_From_Client_Name(ufSaisieHeures.txtClient.value)
        
'        Call Log_Record("     modTEC_Saisie:TEC_Ajoute_Ligne (0024) : ufSaisieHeures.txtDate.Value = '" & ufSaisieHeures.txtDate.value & "' de type " & TypeName(ufSaisieHeures.txtDate.value), -1)
        Dim Y As Integer, m As Integer, d As Integer
        Dim avant As String, apres As String
        On Error Resume Next
            avant = ufSaisieHeures.txtDate.value
            Call Log_Saisie_Heures("info     ", "@00031 - avant = " & avant & "   type = " & TypeName(ufSaisieHeures.txtDate.value) & "   après assignation")
            Y = year(ufSaisieHeures.txtDate.value)
            m = month(ufSaisieHeures.txtDate.value)
            d = day(ufSaisieHeures.txtDate.value)
            If m = 10 Then
                Call Log_Saisie_Heures("info     ", "@00036 - ufSaisieHeures.txtDate.value --->   y = " & Y & "   m = " & m & "   d = " & d & "   type = " & TypeName(ufSaisieHeures.txtDate.value))
            Else
                Call Log_Saisie_Heures("info     ", "@00038  - ufSaisieHeures.txtDate.value - m <> 10 (Erreur) ---->   y = " & Y & "   m = " & m & "    d = " & d & "   type = " & TypeName(ufSaisieHeures.txtDate.value))
            End If
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
        
        Call TEC_Record_Add_Or_Update_To_DB(0) 'Write to MASTER.xlsx file - 2023-12-23 @ 07:03
        Call TEC_Record_Add_Or_Update_Locally(0) 'Write to local worksheet - 2024-02-25 @ 10:34
        
        'Clear the userForm fields after saving
        With ufSaisieHeures
            .txtTEC_ID.value = ""
            .txtClient.value = ""
            .txtActivite.value = ""
            .txtHeures.value = ""
            .txtCommNote.value = ""
            .chbFacturable = True
            .txtSavedHeures.value = ""
        End With
        
        'Code de client
        ufSaisieHeures.txtClient_ID.value = ""
       
        Call TEC_Get_All_TEC_AF
        Call TEC_Refresh_ListBox_And_Add_Hours
        
        Call TEC_TdB_Update_All
        
        'Reset command buttons
        Call Buttons_Enabled_True_Or_False(False, False, False, False)
        
        'Back to client
        ufSaisieHeures.txtClient.SetFocus
    End If
    
    Call Log_Record("modTEC_Saisie:TEC_Ajoute_Ligne", startTime)

End Sub

Sub TEC_Modifie_Ligne() '2023-12-23 @ 07:04

    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_Modifie_Ligne @00086", True)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Modifie_Ligne", 0)

    If Fn_TEC_Is_Data_Valid() = False Then Exit Sub

    'Get the Client_ID
    ufSaisieHeures.txtClient_ID.value = Fn_GetID_From_Client_Name(ufSaisieHeures.txtClient.value)
    
    Call TEC_Record_Add_Or_Update_To_DB(ufSaisieHeures.txtTEC_ID.value)  'Write to external XLSX file - 2023-12-16 @ 14:10
    Call TEC_Record_Add_Or_Update_Locally(ufSaisieHeures.txtTEC_ID.value)  'Write to local worksheet - 2024-02-25 @ 10:38
 
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

    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_Efface_Ligne @00120", True)
    
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

Sub TEC_Get_All_TEC_AF() '2024-09-04 @ 08:47
    
    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_Get_All_TEC_AF @00190", True)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Get_All_TEC_AF - " & _
                                        ufSaisieHeures.txtProf_ID.value & "/" & ufSaisieHeures.txtDate.value, 0)

    Application.ScreenUpdating = False

    'ProfID and Date are mandatory to execute this routine
    If ufSaisieHeures.txtProf_ID.value = "" Or ufSaisieHeures.txtDate.value = "" Then
        Exit Sub
    End If
    
    'Set criteria in worksheet for AdvancedFilter
    wshTEC_Local.Range("R3").value = ufSaisieHeures.txtProf_ID.value
    Call Log_Saisie_Heures("info     ", "@00203 - wshAdmin.Range('TEC_Date').value = " & _
                           ufSaisieHeures.txtDate.value & "   type = " & TypeName(ufSaisieHeures.txtDate.value))
    wshTEC_Local.Range("S3").value = CLng(CDate(ufSaisieHeures.txtDate.value))
    Call Log_Saisie_Heures("info     ", "@00206 - wshTEC_Local.Range('S3').value = " & _
                           wshTEC_Local.Range("S3").value & "   type = " & TypeName(wshTEC_Local.Range("S3").value) & " après assignation")
    wshTEC_Local.Range("T3").value = "FAUX"
    
    With wshTEC_Local
        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
        lastRow = .Cells(.rows.count, "A").End(xlUp).Row 'Last wshTEC_Local Used Row
        If lastRow < 3 Then Exit Sub 'Nothing to filter
        
        'Data Source
        Dim sRng As Range: Set sRng = .Range("A2:P" & lastRow)
        .Range("S10").value = sRng.Address '2024-09-04 @ 08:47
        
         'Criteria
        Dim cRng As Range: Set cRng = .Range("R2:T3")
        .Range("S11").value = cRng.Address '2024-09-04 @ 08:47
        
        'Destination
        Dim dRng As Range: Set dRng = .Range("V2:AI2")
        .Range("S12").value = dRng.Address '2024-09-04 @ 08:47
        
        'Advanced Filter applied to BaseHours (Prof, Date and isDetruit)
        sRng.AdvancedFilter xlFilterCopy, cRng, dRng, False
        
        lastResultRow = .Cells(.rows.count, "V").End(xlUp).Row
        .Range("S13").value = (lastResultRow - 2) & " rows" '2024-09-04 @ 08:47
        .Range("S14").value = Format$(Now(), "yyyy-mm-dd hh:mm:ss") '2024-09-04 @ 08:47
        
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort 'Sort - Date / Prof / TEC_ID
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
    
    'Libérer la mémoire
    Set cRng = Nothing
    Set dRng = Nothing
    Set sRng = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Get_All_TEC_AF", startTime)

End Sub

Sub TEC_Efface_Formulaire() 'Clear all fields on the userForm

    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_Efface_Formulaire @00279", True)
    
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
    
    Call Buttons_Enabled_True_Or_False(False, False, False, False)
        
    ufSaisieHeures.txtClient.SetFocus
    
    Call Log_Record("modTEC_Saisie:TEC_Efface_Formulaire", startTime)

End Sub

Sub TEC_Record_Add_Or_Update_To_DB(TECID As Long) 'Write -OR- Update a record to external .xlsx file
    
    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB @00309", True)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB(" & TECID & ")", 0)

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
    
    Dim dateValue As Date '2024-09-04 @ 09:01
    dateValue = ufSaisieHeures.txtDate.value
    'Special log to debug Date Format issue... 2024-09-06 @ 16:32
    If month(dateValue) < 10 Then
        Call Log_Saisie_Heures("Erreur   ", "00338 - E R R E U R (mois < 10) - dateValue = " & dateValue & "  type = " & TypeName(dateValue))
    End If
    If TECID = 0 And Now - dateValue > 15 Then
        MsgBox "La date saisie est plus de 15 jours dans le passé..." & vbNewLine & vbNewLine & _
                "Veuillez aviser le développeur de cette situation SVP", vbInformation
        Call Log_Saisie_Heures("Future   ", "Plus de 15 jours dans le passé - dateValue = " & dateValue & "  type = " & TypeName(dateValue))
    End If
    
    If TECID < 0 Then 'Soft delete a record
        
        'Open the recordset for the specified ID
        
        Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB - SoftDeleteStart - " & CStr(saveLogTEC_ID), -1) '2024-09-13 @ 08:35
        
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & Abs(TECID), conn, 2, 3
        saveLogTEC_ID = TECID
        If Not rs.EOF Then
            'Update the "IsDeleted" field to mark the record as deleted
            rs.Fields("DateSaisie").value = CDate(Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
            rs.Fields("EstDetruit").value = ConvertValueBooleanToText(True)
            rs.Fields("VersionApp").value = ThisWorkbook.name
            rs.update
            
            Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB - " & CStr(-(saveLogTEC_ID)) & " I S   D E L E T E D", -1) '2024-09-13 @ 08:49
        
            Call Log_Saisie_Heures("DELETE" & saveLogTEC_ID, ufSaisieHeures.cmbProfessionnel.value & " | " & _
                                    dateValue & " | " & _
                                    ufSaisieHeures.txtClient_ID.value & " | " & _
                                    ufSaisieHeures.txtClient.value & " | " & _
                                    ufSaisieHeures.txtActivite.value & " | " & _
                                    Format$(ufSaisieHeures.txtHeures.value, "#0.00") & " | " & _
                                    ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value) & " | " & _
                                    ufSaisieHeures.txtCommNote.value)

        Else 'Handle the case where the specified ID is not found - PROBLEM !!!
            
            MsgBox "L'enregistrement avec le TEC_ID '" & TECID & "' ne peut être trouvé!", _
                vbExclamation
                
            Call Log_Record("N'a pas trouvé le TECID - '" & CStr(saveLogTEC_ID) & "'", -1) '2024-09-02 @ 10:35
            
            rs.Close
            conn.Close
            
            Exit Sub
        End If
    
    Else 'Add a new record (TECID = 0) -OR- update an existing one (TECID <> 0)
        
        If TECID = 0 Then 'Add a record
        
            Call Log_Record("     modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB - AddingNewRecord - TECID = " & CStr(saveLogTEC_ID), -1) '2024-09-13 @ 08:35
        
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
            
            Call Log_Record("     modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB - Le nouvel enregistrement, TECID a été assigné à " & CStr(nextID), -1) '2024-09-13 @ 08:35

            ufSaisieHeures.txtTEC_ID.value = nextID
            saveLogTEC_ID = nextID
        
            'Close the previous recordset, no longer needed and open an empty recordset
            rs.Close
            rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
            
            Call Log_Record("     modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB - Création du RecordSet", -1) '2024-09-13 @ 08:35

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
            rs.Fields("EstFacturable").value = ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value)
            rs.Fields("DateSaisie").value = CDate(Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
            rs.Fields("EstFacturee").value = ConvertValueBooleanToText(False)
            rs.Fields("DateFacturee").value = Null
            rs.Fields("EstDetruit").value = ConvertValueBooleanToText(False)
            rs.Fields("VersionApp").value = ThisWorkbook.name
            rs.Fields("NoFacture").value = ""
            rs.update
            
            Call Log_Record("     modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB - Le RecordSet a été ajouté pour le TECID = " & CStr(nextID), -1) '2024-09-13 @ 08:35
            
            'Nouveau log - 2024-09-02 @ 10:40
            Call Log_Saisie_Heures("ADD    " & saveLogTEC_ID, ufSaisieHeures.cmbProfessionnel.value & " | " & _
                        dateValue & " | " & _
                        ufSaisieHeures.txtClient_ID.value & " | " & _
                        ufSaisieHeures.txtClient.value & " | " & _
                        ufSaisieHeures.txtActivite.value & " | " & _
                        Format$(ufSaisieHeures.txtHeures.value, "#0.00") & " | " & _
                        ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value) & " | " & _
                        ufSaisieHeures.txtCommNote.value)
        
        Else 'Update an existing record (TECID <> 0)
        
            'Open the recordset for the specified ID
            rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & TECID, conn, 2, 3
            If Not rs.EOF Then
                Call Log_Record("      modTEC_Saisie" & "TEC_Record_Add_Or_Update_To_DB" & " - Update an existing record - Create a new RecordSet " & CStr(nextID))  '2024-09-13 @ 08:35
                'Update fields for the existing record
                rs.Fields("Client_ID").value = ufSaisieHeures.txtClient_ID.value
                rs.Fields("ClientNom").value = ufSaisieHeures.txtClient.value
                rs.Fields("Description").value = ufSaisieHeures.txtActivite.value
                rs.Fields("Heures").value = Format$(ufSaisieHeures.txtHeures.value, "#0.00")
                rs.Fields("CommentaireNote").value = ufSaisieHeures.txtCommNote.value
                rs.Fields("EstFacturable").value = ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value)
                rs.Fields("DateSaisie").value = CDate(Format$(Now(), "yyyy/mm/dd hh:mm:ss"))
                rs.Fields("VersionApp").value = ThisWorkbook.name
                'Nouveau log - 2024-09-02 @ 10:40
                
                Call Log_Saisie_Heures("UPDATE " & saveLogTEC_ID, ufSaisieHeures.cmbProfessionnel.value & " | " & _
                            dateValue & " | " & _
                            ufSaisieHeures.txtClient_ID.value & " | " & _
                            ufSaisieHeures.txtClient.value & " | " & _
                            ufSaisieHeures.txtActivite.value & " | " & _
                            Format$(ufSaisieHeures.txtHeures.value, "#0.00") & " | " & _
                            ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value) & " | " & _
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
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_To_DB - " & TECID, startTime)
    
End Sub

Sub TEC_Record_Add_Or_Update_Locally(TECID As Long) 'Write -OR- Update a record to local worksheet
    
    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_Record_Add_Or_Update_Locally @00536", True)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_Locally(" & TECID & ")", 0)

    Application.ScreenUpdating = False
    
    'What is the row number of this TEC_ID ?
    Dim lastUsedRow As Long
    
    Dim hoursValue As Double '2024-03-01 @ 05:40
    hoursValue = CDbl(ufSaisieHeures.txtHeures.value)
    
    Dim dateValue As Date
    dateValue = ufSaisieHeures.txtDate.value
    Call Log_Saisie_Heures("info     ", "@00550 - dateValue = " & dateValue & "   type = " & TypeName(dateValue) & "   après assignation")
    
    If TECID = 0 Then 'Add a new record
        Call Log_Record("     modTEC_Saisie:TEC_Record_Add_Or_Update_Locally - Add new record", -1) '2024-10-05 @ 07:32
        'Get the next available row in TEC_Local
        Dim nextRowNumber As Long
        nextRowNumber = wshTEC_Local.Range("A9999").End(xlUp).Row + 1
        With wshTEC_Local
            .Range("A" & nextRowNumber).value = ufSaisieHeures.txtTEC_ID.value
            .Range("B" & nextRowNumber).value = ufSaisieHeures.txtProf_ID.value
            .Range("C" & nextRowNumber).value = ufSaisieHeures.cmbProfessionnel.value
            .Range("D" & nextRowNumber).value = dateValue
            .Range("E" & nextRowNumber).value = ufSaisieHeures.txtClient_ID.value
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
        Call Log_Record("     modTEC_Saisie:TEC_Record_Add_Or_Update_Locally - Update -OR- Delete - " & CStr(TECID), -1)
        'What is the row number for the TEC_ID
        lastUsedRow = wshTEC_Local.Range("A99999").End(xlUp).Row
        Dim lookupRange As Range:  Set lookupRange = wshTEC_Local.Range("A3:A" & lastUsedRow)
        Dim rowToBeUpdated As Long
        rowToBeUpdated = Fn_Find_Row_Number_TEC_ID(Abs(TECID), lookupRange)
        If rowToBeUpdated < 1 Then
            'Handle the case where the specified TecID is not found !!
            MsgBox "L'enregistrement avec le TEC_ID '" & TECID & "' ne peut être trouvé!", _
                vbExclamation
            Exit Sub
        End If

        If TECID > 0 Then 'Modify the record
            Call Log_Record("     modTEC_Saisie:TEC_Record_Add_Or_Update_Locally - Update - " & CStr(TECID), -1)
            With wshTEC_Local
                .Range("E" & rowToBeUpdated).value = ufSaisieHeures.txtClient_ID.value
                .Range("F" & rowToBeUpdated).value = ufSaisieHeures.txtClient.value
                .Range("G" & rowToBeUpdated).value = ufSaisieHeures.txtActivite.value
                .Range("H" & rowToBeUpdated).value = hoursValue
                .Range("I" & rowToBeUpdated).value = ufSaisieHeures.txtCommNote.value
                .Range("J" & rowToBeUpdated).value = ConvertValueBooleanToText(ufSaisieHeures.chbFacturable.value)
                .Range("K" & rowToBeUpdated).value = CDate(Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
                .Range("L" & rowToBeUpdated).value = ConvertValueBooleanToText(False)
                .Range("M" & rowToBeUpdated).value = ""
                .Range("N" & rowToBeUpdated).value = ConvertValueBooleanToText(False)
                .Range("O" & rowToBeUpdated).value = ThisWorkbook.name
                .Range("P" & rowToBeUpdated).value = ""
            End With
        Else 'Soft delete the record
            Call Log_Record("     modTEC_Saisie:TEC_Record_Add_Or_Update_Locally - SoftDelete - " & CStr(-TECID), -1)
            With wshTEC_Local
                .Range("K" & rowToBeUpdated).value = CDate(Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
                .Range("N" & rowToBeUpdated).value = ConvertValueBooleanToText(True)
                .Range("O" & rowToBeUpdated).value = ThisWorkbook.name
            End With
        End If
    End If
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Record_Add_Or_Update_Locally - " & TECID, startTime)

End Sub

Sub TEC_Refresh_ListBox_And_Add_Hours() 'Load the listBox with the appropriate records

    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_Refresh_ListBox_And_Add_Hours @00621", True)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_Refresh_ListBox_And_Add_Hours", 0)

    If ufSaisieHeures.txtProf_ID.value = "" Or ufSaisieHeures.txtDate.value = "" Then
        GoTo EndOfProcedure
    End If
    
    'Modifie le critère pour forcer une execution du AdvancedFilter
    wshTEC_TDB_Data.Range("S6").value = ufSaisieHeures.cmbProfessionnel.value
    
    ufSaisieHeures.txtTotalHeures.value = ""
    ufSaisieHeures.txtHresFact.value = ""
    ufSaisieHeures.txtHresNF.value = ""
    ufSaisieHeures.txtHresFactSemaine.value = ""
    ufSaisieHeures.txtHresNFSemaine.value = ""

    ufSaisieHeures.lsbHresJour.RowSource = ""
    ufSaisieHeures.lsbHresJour.Clear '2024-08-10 @ 05:59
    
    'Last Row used in first column of result
    Dim lastRow As Long
    lastRow = wshTEC_Local.Range("V999").End(xlUp).Row
    If lastRow < 3 Then GoTo Rien_Aujourdhui
        
    With ufSaisieHeures.lsbHresJour
        .ColumnHeads = False
        .ColumnCount = 9
        .ColumnWidths = "30; 24; 54; 155; 240; 35; 90; 32; 90"
'        .RowSource = wshTEC_Local.name & "!V3:AI" & lastRow '2024-08-11 @ 12:50
    End With
    
    'Manually add to listBox (.RowSource DOES NOT WORK!!!)
    Dim rng As Range
    Set rng = wshTEC_Local.Range("V3:AI" & lastRow)
'    Debug.Print rng.Address
     
    Dim i As Long, j As Long
    Dim totalHeures As Double
    Dim totalHresFact As Double, totalHresNonFact As Double
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
        If Fn_Is_Client_Facturable(rng.Cells(i, 14).value) = True And rng.Cells(i, 8).value = "VRAI" Then
            totalHresFact = totalHresFact + CCur(rng.Cells(i, 6).value)
        Else
            totalHresNonFact = totalHresNonFact + CCur(rng.Cells(i, 6).value)
        End If
    Next i
         
Rien_Aujourdhui:

    ufSaisieHeures.txtTotalHeures.value = Format$(totalHeures, "#0.00")
    ufSaisieHeures.txtHresFact.value = Format$(totalHresFact, "#0.00")
    ufSaisieHeures.txtHresNF.value = Format$(totalHresNonFact, "#0.00")
    
    'Additionner les charges de la semaine
    Dim totalHresFactSemaine As Double
    Dim totalHresNonFactSemaine As Double
    
    lastRow = wshTEC_TDB_Data.Cells(wshTEC_TDB_Data.rows.count, "W").End(xlUp).Row
    For i = 2 To lastRow
        totalHresFactSemaine = totalHresFactSemaine + wshTEC_TDB_Data.Range("AC" & i).value
        totalHresNonFactSemaine = totalHresNonFactSemaine + wshTEC_TDB_Data.Range("AD" & i).value
    Next i
    
    ufSaisieHeures.txtHresFactSemaine.value = Format$(totalHresFactSemaine, "#0.00")
    ufSaisieHeures.txtHresNFSemaine.value = Format$(totalHresNonFactSemaine, "#0.00")
    
    ufSaisieHeures.Repaint
    
    DoEvents '2024-08-12 @ 10:31
    
    Application.ScreenUpdating = True

EndOfProcedure:

    Call Buttons_Enabled_True_Or_False(False, False, False, False)

    ufSaisieHeures.txtClient.SetFocus
    
    'Libérer la mémoire
    Set rng = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_Refresh_ListBox_And_Add_Hours", startTime)
    
End Sub

Sub TEC_TdB_Push_TEC_Local_To_DB_Data()

    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_TdB_Push_TEC_Local_To_DB_Data @00686", True)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_TdB_Push_TEC_Local_To_DB_Data", 0)

    Dim wsFrom As Worksheet: Set wsFrom = wshTEC_Local
    
    Dim lastUsedRow As Long
    lastUsedRow = wshTEC_Local.Range("A99999").End(xlUp).Row
    
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

    Dim rngTo As Range: Set rngTo = wshTEC_TDB_Data.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2))
    rngTo.value = arr
    
    'Libérer la mémoire
    Set rngTo = Nothing
    Set wsFrom = Nothing
    
    Call Log_Record("modTEC_Saisie:TEC_TdB_Push_TEC_Local_To_DB_Data", startTime)

End Sub

Sub TEC_TdB_Update_All()

    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_TdB_Update_All @00726", True)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_TdB_Update_All", 0)
    
    Call TEC_TdB_Push_TEC_Local_To_DB_Data
    
    Call TEC_TdB_Refresh_All_Pivot_Tables
    
    Call Log_Record("modTEC_Saisie:TEC_TdB_Update_All", startTime)

End Sub

Sub TEC_TdB_Refresh_All_Pivot_Tables()

    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:TEC_TdB_Refresh_All_Pivot_Tables @00739", True)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Saisie:TEC_TdB_Refresh_All_Pivot_Tables", 0)

    Dim pt As pivotTable
    For Each pt In wshTEC_TDB_PivotTable.PivotTables
        pt.RefreshTable
    Next pt

    Call Log_Record("modTEC_Saisie:TEC_TdB_Refresh_All_Pivot_Tables", startTime)

    'Libérer la mémoire
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
        filterDate = dateValue("02/09/2024")
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
                .SortFields.Clear
                .SortFields.Add key:=wshTEC_Local.Range("AT3"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Sort Based On Date
                .SortFields.Add key:=wshTEC_Local.Range("AR3"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Sort Based On Prof_ID
                .SortFields.Add key:=wshTEC_Local.Range("AQ3"), _
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
    'Libérer la mémoire
    Set cRng = Nothing
    Set dRng = Nothing
    Set sRng = Nothing
    Set ws = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite lors de l'application du filtre avancé.", vbCritical
    On Error GoTo 0
    Resume Cleaning

End Sub

Sub Buttons_Enabled_True_Or_False(a As Boolean, u As Boolean, _
                                  d As Boolean, c As Boolean)
                                  
    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:Buttons_Enabled_True_Or_False @00833", True)
    
    Call Log_Saisie_Heures("info     ", "@00835 - Paramètres : add = " & a & "  update = " & u & "  delete = " & d & "  clear = " & c)
    
    With ufSaisieHeures
        .cmdAdd.Enabled = a
        .cmdUpdate.Enabled = u
        .cmdDelete.Enabled = d
        .cmdClear.Enabled = c
    End With

End Sub

Sub MsgBoxInvalidDate(location As String) '2024-06-13 @ 12:40

    Call Log_Saisie_Heures("entering ", "E n t e r i n g   modTEC_Saisie:MsgBoxInvalidDate @00848", True)
    
    MsgBox "La date saisie ne peut être acceptée tel qu'elle est entrée." & vbNewLine & vbNewLine & _
           "Elle doit être obligatoirement de format:" & vbNewLine & _
           "     'j', jj', " & vbNewLine & _
           "     'jj-mm', 'jj/mm' ou " & vbNewLine & _
           "     'j-m-aa', 'j-m-aaaa', 'jj-mm-aaaa'" & vbNewLine & vbNewLine & _
           "Veuillez saisir la date de nouveau SVP", _
           vbCritical, _
           "La date saisie est INVALIDE - " & location

End Sub


