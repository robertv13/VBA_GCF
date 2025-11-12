Attribute VB_Name = "modTEC_Saisie"
Option Explicit

Public Const MODE_INITIAL_FAC As Long = 1
Public Const MODE_CREATION_FAC As Long = 2
Public Const MODE_AFFICHAGE_FAC As Long = 3
Public Const MODE_MODIFICATION_FAC As Long = 4

Public rmv_state As Long

Sub AjouterLigneTEC()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:AjouterLigneTEC", vbNullString, 0)

    Dim timerPerformance As Double: timerPerformance = Timer
    
    'Obtenir le ID du client pur (à partir de son nom pur)
    ufSaisieHeures.txtClientID.Value = Fn_CellSpecifiqueDeBDClient(ufSaisieHeures.txtClient.Value, 1, 2)
        
    If Fn_TEC_Is_Data_Valid() = True Then
        Dim Y As Integer, m As Integer, d As Integer
        On Error Resume Next
            Y = year(ufSaisieHeures.txtDate.Value)
            m = month(ufSaisieHeures.txtDate.Value)
            d = day(ufSaisieHeures.txtDate.Value)
            ufSaisieHeures.txtDate.Value = Format$(DateSerial(Y, m, d), "yyyy-mm-dd")
        On Error GoTo 0
        
        If Fn_Is_Client_Facturable(ufSaisieHeures.txtClientID) = False Then '2025-10-31 @ 08:30
            ufSaisieHeures.chkFacturable = False
        End If
        
        Call AjouterTECdansBDMaster
        Call AjouterOuModifierTECdansBDLocale(0)
        
        Dim saveTECID As String '2025-11-05 @ 08:41
        saveTECID = CStr(ufSaisieHeures.txtTECID.Value)
    
        'Clear the userForm fields after saving
        With ufSaisieHeures
            .txtTECID.Value = vbNullString
            .txtClient.Value = vbNullString
            .txtClientID.Value = vbNullString
            .txtActivite.Value = vbNullString
            .txtHeures.Value = vbNullString
            .txtCommNote.Value = vbNullString
            .chkFacturable = True
        End With
        
        ufSaisieHeures.valeurSauveeHeures = 0 '2025-05-07 @ 16:54
        
        Call ObtenirTousLesTECDateAvecAF
        Call RafraichirListBoxEtAddtionnerHeures
        
        'Back to client
        ufSaisieHeures.txtClient.SetFocus
    Else
        Exit Sub
    End If
    
    Call EnregistrerLogPerformance("modTEC_Saisie.AjouterLigneTEC - " & saveTECID, _
                                        Timer - timerPerformance) '2025-11-05 @ 08:41
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:AjouterLigneTEC", vbNullString, startTime)

End Sub

Sub ModifierLigneTEC() '2023-12-23 @ 07:04

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:ModifierLigneTEC", vbNullString, 0)

    If Fn_TEC_Is_Data_Valid() = False Then Exit Sub

    Dim timerPerformance As Double: timerPerformance = Timer
        
    'Obtenir le ID du client pur (à partir de son nom pur) - 2025-03-04 @ 08:02
    ufSaisieHeures.txtClientID.Value = Fn_CellSpecifiqueDeBDClient(ufSaisieHeures.txtClient.Value, 1, 2)
        
    If Fn_Is_Client_Facturable(ufSaisieHeures.txtClientID) = False Then '2025-10-31 @ 08:30
        ufSaisieHeures.chkFacturable = False
    End If
    
    Call ModifierTECdansBDMaster(ufSaisieHeures.txtTECID.Value) '2025-10-31 @ 06:43
    Call AjouterOuModifierTECdansBDLocale(ufSaisieHeures.txtTECID.Value)
 
    Dim saveTECID As String '2025-11-05 @ 08:41
    saveTECID = CStr(ufSaisieHeures.txtTECID.Value)
    
    'Initialize dynamic variables
    With ufSaisieHeures
        .txtTECID.Value = vbNullString
        .cmbProfessionnel.Enabled = True
        .txtDate.Enabled = True
        .txtClient.Value = vbNullString
        .txtActivite.Value = vbNullString
        .txtHeures.Value = vbNullString
        .txtCommNote.Value = vbNullString
        .chkFacturable = True
    End With

    Call ObtenirTousLesTECDateAvecAF
    Call RafraichirListBoxEtAddtionnerHeures
    
    rmv_state = MODE_CREATION_FAC
    
    ufSaisieHeures.txtClient.SetFocus
    
    Call EnregistrerLogPerformance("modTEC_Saisie.ModifierLigneTEC - " & saveTECID, _
                                        Timer - timerPerformance) '2025-11-05 @ 08:41
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:ModifierLigneTEC", vbNullString, startTime)

End Sub

Sub DetruireLigneTEC() '2023-12-23 @ 07:05

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:DetruireLigneTEC", vbNullString, 0)

    If ufSaisieHeures.txtTECID.Value = vbNullString Then
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
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:DetruireLigneTEC - Le DELETE est confirmé - ", _
                CStr(-ufSaisieHeures.txtTECID.Value), -1)
    
    Dim timerPerformance As Double: timerPerformance = Timer
        
    Dim Sh As Worksheet: Set Sh = wsdTEC_Local
    
    Dim tecID As Long
    tecID = -ufSaisieHeures.txtTECID.Value
    
    If Fn_Is_Client_Facturable(ufSaisieHeures.txtClientID) = False Then '2025-10-31 @ 08:30
        ufSaisieHeures.chkFacturable = False
    End If
    
    Call SupprimerTECdansBDMaster(tecID) '2025-10-31@ 06:44
    Call AjouterOuModifierTECdansBDLocale(tecID)
    
    'Empty the dynamic fields after deleting
    With ufSaisieHeures
        .txtClient.Value = vbNullString
        .txtActivite.Value = vbNullString
        .txtHeures.Value = vbNullString
        .txtCommNote.Value = vbNullString
        .chkFacturable = True
    End With
    
    Call EnregistrerLogPerformance("modTEC_Saisie.DetruireLigneTEC - " & CStr(Abs(tecID)), _
                                        Timer - timerPerformance) '2025-11-05 @ 08:41
    MsgBox _
        Prompt:="L'enregistrement a été DÉTRUIT !", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        
    ufSaisieHeures.cmbProfessionnel.Enabled = True
    ufSaisieHeures.txtDate.Enabled = True
    rmv_state = MODE_CREATION_FAC
    
    Call ObtenirTousLesTECDateAvecAF
    Call RafraichirListBoxEtAddtionnerHeures
    
Clean_Exit:

    ufSaisieHeures.txtTECID.Value = vbNullString
    ufSaisieHeures.txtClient.SetFocus

    'Libérer la mémoire
    Set Sh = Nothing

    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:DetruireLigneTEC", vbNullString, startTime)

End Sub

Sub ObtenirTousLesTECDateAvecAF()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:ObtenirTousLesTECDateAvecAF", _
                                                                 ufSaisieHeures.txtProfID.Value & " / " & ufSaisieHeures.txtDate.Value, 0)

    Dim ws As Worksheet: Set ws = wsdTEC_Local
    
    Application.ScreenUpdating = False

    'ProfID and Date are mandatory to execute this routine
    If ufSaisieHeures.txtProfID.Value = vbNullString Or ufSaisieHeures.txtDate.Value = vbNullString Then
        Exit Sub
    End If
    
    'wsdTEC_Local_AF#1

    'Set criteria directly in TEC_Local for AdvancedFilter
    With ws
        .Range("R3").Value = ufSaisieHeures.txtProfID.Value
        .Range("S3").Value = CLng(CDate(ufSaisieHeures.txtDate.Value))
        .Range("T3").Value = "FAUX"
    End With
    
    'Effacer les données de la dernière utilisation
    ws.Range("S6:S10").ClearContents
    ws.Range("S6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_TEC_Local[#All]")
    ws.Range("S7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("R2:T3")
    ws.Range("S8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("V1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("V2:AI2")
    ws.Range("S9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False

    Dim lastResultRow As Long
    lastResultRow = ws.Cells(ws.Rows.count, "V").End(xlUp).Row
    ws.Range("S10").Value = (lastResultRow - 2) & " lignes"
        
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
        Call modAppli_Utils.ConvertirPlageABooleen(r)
        Set r = wsdTEC_Local.Range("AE3:AE" & lastResultRow)
        Call modAppli_Utils.ConvertirPlageABooleen(r)
        Set r = wsdTEC_Local.Range("AG3:AG" & lastResultRow)
        Call modAppli_Utils.ConvertirPlageABooleen(r)
    End If
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set r = Nothing
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:ObtenirTousLesTECDateAvecAF", vbNullString, startTime)

End Sub

Sub EffacerFormulaireTEC() '2025-07-03 @ 07:31

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:EffacerFormulaireTEC", vbNullString, 0)

    'Empty the dynamic fields after reseting the form
    With ufSaisieHeures
        .txtTECID.Value = vbNullString
        .txtClient.Value = vbNullString
        .txtClientID.Value = vbNullString
        .txtActivite.Value = vbNullString
        .txtHeures.Value = vbNullString
        .txtCommNote.Value = vbNullString
        .cmbProfessionnel.Enabled = True
        .txtDate.Enabled = True
    End With
    
    Call ObtenirTousLesTECDateAvecAF
    Call RafraichirListBoxEtAddtionnerHeures
    
    ufSaisieHeures.txtClient.SetFocus
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:EffacerFormulaireTEC", vbNullString, startTime)

End Sub

Public Function AjouterTECdansBDMaster() As Boolean '2025-10-31 @ 06:03

    On Error GoTo GestionErreur
    AjouterTECdansBDMaster = False
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:AjouterTECdansBDMaster", vbNullString, 0)

    Dim conn As Object, recSet As Object
    Set conn = CreateObject("ADODB.Connection")
    Set recSet = CreateObject("ADODB.Recordset")

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "TEC_Local$"
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Obtenir le prochain TECID
    Dim rsID As Object, nextID As Long
    Set rsID = CreateObject("ADODB.Recordset")
    rsID.Open "SELECT MAX(TECID) AS MaxID FROM [" & destinationTab & "]", conn
    nextID = IIf(IsNull(rsID.Fields("MaxID").Value), 1, rsID.Fields("MaxID").Value + 1)
    rsID.Close: Set rsID = Nothing

    ufSaisieHeures.txtTECID.Value = nextID

    'Ajouter l'enregistrement
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    recSet.AddNew
    Call RemplirChampsRecordset(recSet, nextID, ufSaisieHeures.txtDate.Value, Now, False)
    recSet.Update

    Call EnregistrerLogSaisieHeures("ADD    " & nextID, ConstruireLigneLog())
    AjouterTECdansBDMaster = True

Nettoyage:
    On Error Resume Next
    recSet.Close: conn.Close
    Set recSet = Nothing: Set conn = Nothing
'    If Dir(verrouPath) <> "" Then Kill verrouPath
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:AjouterTECdansBDMaster", CStr(nextID), startTime)
    
    Exit Function

GestionErreur:
    MsgBox "Erreur lors de l'ajout du TEC : " & Err.description, _
            vbCritical
    Resume Nettoyage
    
End Function

Public Function ModifierTECdansBDMaster(tecID As Long) As Boolean '2025-10-31 @06:05

    On Error GoTo GestionErreur
    ModifierTECdansBDMaster = False

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:ModifierTECdansBDMaster", vbNullString, 0)
    
    Dim conn As Object, recSet As Object
    Set conn = CreateObject("ADODB.Connection")
    Set recSet = CreateObject("ADODB.Recordset")

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "TEC_Local$"
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & tecID, conn, 2, 3
    If recSet.EOF Then
        MsgBox "TECID introuvable pour modification : " & tecID, vbExclamation
        GoTo Nettoyage
    End If

    Call RemplirChampsRecordset(recSet, tecID, ufSaisieHeures.txtDate.Value, Now, True)
    recSet.Update

    Call EnregistrerLogSaisieHeures("UPDATE " & tecID, ConstruireLigneLog())
    ModifierTECdansBDMaster = True

Nettoyage:
    On Error Resume Next
    recSet.Close: conn.Close
    Set recSet = Nothing: Set conn = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:ModifierTECdansBDMaster", CStr(tecID), startTime)
    
    Exit Function

GestionErreur:
    MsgBox "Erreur lors de la modification du TEC : " & Err.description, vbCritical
    Resume Nettoyage
    
End Function

Public Function SupprimerTECdansBDMaster(tecID As Long) As Boolean '2025-10-31 @ 06:06

    On Error GoTo GestionErreur
    SupprimerTECdansBDMaster = False

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:SupprimerTECdansBDMaster", vbNullString, 0)
    
    Dim conn As Object, recSet As Object
    Set conn = CreateObject("ADODB.Connection")
    Set recSet = CreateObject("ADODB.Recordset")

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "TEC_Local$"
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & Abs(tecID), conn, 2, 3
    If recSet.EOF Then
        MsgBox "TECID introuvable pour suppression : " & tecID, vbExclamation
        GoTo Nettoyage
    End If

    With recSet
        .Fields(fTECDateSaisie - 1).Value = Format$(Now, "yyyy-mm-dd hh:mm:ss")
        .Fields(fTECEstDetruit - 1).Value = Fn_Convert_Value_Boolean_To_Text(True)
        .Fields(fTECVersionApp - 1).Value = ThisWorkbook.Name
        .Update
    End With

    Call EnregistrerLogSaisieHeures("DELETE " & Abs(tecID), ConstruireLigneLog())
    SupprimerTECdansBDMaster = True

Nettoyage:
    On Error Resume Next
    recSet.Close: conn.Close
    Set recSet = Nothing: Set conn = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:SupprimerTECdansBDMaster", CStr(Abs(tecID)), startTime)
    
    Exit Function

GestionErreur:
    MsgBox "Erreur lors de la suppression du TEC : " & Err.description, vbCritical
    Resume Nettoyage
    
End Function

Sub AjouterOuModifierTECdansBDLocale(tecID As Long)
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:AjouterOuModifierTECdansBDLocale", CStr(tecID), 0)

    Application.ScreenUpdating = False
    
    'What is the row number of this TECID ?
    Dim lastUsedRow As Long
    
    Dim hoursValue As Double
    hoursValue = CDbl(ufSaisieHeures.txtHeures.Value)
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim dateValue As Date
    dateValue = ufSaisieHeures.txtDate.Value
    
    If tecID = 0 Then 'Add a new record
        'Get the next available row in TEC_Local
        Dim nextRowNumber As Long
        With wsdTEC_Local
            nextRowNumber = .Cells(.Rows.count, 1).End(xlUp).Row + 1
            .Range("A" & nextRowNumber).Value = ufSaisieHeures.txtTECID.Value
            .Range("B" & nextRowNumber).Value = ufSaisieHeures.txtProfID.Value
            .Range("C" & nextRowNumber).Value = ufSaisieHeures.cmbProfessionnel.Value
            .Range("D" & nextRowNumber).Value = dateValue
            .Range("E" & nextRowNumber).Value = ufSaisieHeures.txtClientID.Value
            .Range("F" & nextRowNumber).Value = ufSaisieHeures.txtClient.Value
            .Range("G" & nextRowNumber).Value = ufSaisieHeures.txtActivite.Value
            .Range("H" & nextRowNumber).Value = hoursValue
            .Range("I" & nextRowNumber).Value = ufSaisieHeures.txtCommNote.Value
            .Range("J" & nextRowNumber).Value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chkFacturable.Value)
            .Range("K" & nextRowNumber).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            .Range("L" & nextRowNumber).Value = Fn_Convert_Value_Boolean_To_Text(False)
            .Range("M" & nextRowNumber).Value = vbNullString
            .Range("N" & nextRowNumber).Value = Fn_Convert_Value_Boolean_To_Text(False)
            .Range("O" & nextRowNumber).Value = ThisWorkbook.Name
            .Range("P" & nextRowNumber).Value = vbNullString
        End With
    Else
        'What is the row number for the TECID
        lastUsedRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "A").End(xlUp).Row
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
                .Range("E" & rowToBeUpdated).Value = ufSaisieHeures.txtClientID.Value
                .Range("F" & rowToBeUpdated).Value = ufSaisieHeures.txtClient.Value
                .Range("G" & rowToBeUpdated).Value = ufSaisieHeures.txtActivite.Value
                .Range("H" & rowToBeUpdated).Value = hoursValue
                .Range("I" & rowToBeUpdated).Value = ufSaisieHeures.txtCommNote.Value
                .Range("J" & rowToBeUpdated).Value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chkFacturable.Value)
                .Range("K" & rowToBeUpdated).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
                .Range("L" & rowToBeUpdated).Value = Fn_Convert_Value_Boolean_To_Text(False)
                .Range("M" & rowToBeUpdated).Value = vbNullString
                .Range("N" & rowToBeUpdated).Value = Fn_Convert_Value_Boolean_To_Text(False)
                .Range("O" & rowToBeUpdated).Value = ThisWorkbook.Name
                .Range("P" & rowToBeUpdated).Value = vbNullString
            End With
        Else 'Soft delete the record
            With wsdTEC_Local
                .Range("K" & rowToBeUpdated).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
                .Range("N" & rowToBeUpdated).Value = Fn_Convert_Value_Boolean_To_Text(True)
                .Range("O" & rowToBeUpdated).Value = ThisWorkbook.Name
            End With
        End If
    End If
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:AjouterOuModifierTECdansBDLocale", CStr(tecID), startTime)

End Sub

Sub RafraichirListBoxEtAddtionnerHeures() 'Load the listBox with the appropriate records

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:RafraichirListBoxEtAddtionnerHeures", _
            ufSaisieHeures.txtProfID.Value & " / " & ufSaisieHeures.txtDate.Value, 0)

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If ufSaisieHeures.txtProfID.Value = vbNullString Or Not IsDate(ufSaisieHeures.txtDate.Value) Then
        MsgBox "Veuillez entrer un professionnel et/ou une date valide.", vbExclamation
        GoTo EndOfProcedure
    End If
    
    'On vide le formulaire
    ufSaisieHeures.txtTotalHeures.Value = vbNullString
    ufSaisieHeures.txtHresFact.Value = vbNullString
    ufSaisieHeures.txtHresNF.Value = vbNullString
    ufSaisieHeures.txtHresFactSemaine.Value = vbNullString
    ufSaisieHeures.txtHresNFSemaine.Value = vbNullString

    ufSaisieHeures.lstHresJour.RowSource = vbNullString
    ufSaisieHeures.lstHresJour.Clear
    
    With ufSaisieHeures.lstHresJour
        .ColumnHeads = False
        .ColumnCount = 9
        .ColumnWidths = "30; 23; 60; 157; 242; 35; 90; 32; 90"
    End With
    
    'Manually add to listBox (because some tests have to be made)
    Dim lastRow As Long
    lastRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "V").End(xlUp).Row
    Dim data As Variant
    data = wsdTEC_Local.Range("V3:AI" & lastRow).Value
     
    'Variables initiales
    Dim totalHeures As Currency: totalHeures = 0
    Dim totalHresFact As Currency: totalHresFact = 0
    Dim totalHresNonFact As Currency: totalHresNonFact = 0
    Dim rngResult As Range
    Dim i As Long, ColIndex As Long
    
    'Remplissage du listBox '2025-07-03 @ 08:37
    Dim hresFormat As String
    Dim nbLignes As Long: nbLignes = UBound(data, 1)
    Dim nbColonnes As Long: nbColonnes = UBound(data, 2)
    
    Dim totaux As Collection
    Set totaux = New Collection
    totaux.Add 0
    totaux.Add 0
    totaux.Add 0
    
    If lastRow >= 3 Then
'        Set rng = wsdTEC_Local.Range("V3:AI" & lastRow)
        For i = 1 To nbLignes
            With ufSaisieHeures.lstHresJour
                .AddItem data(i, 1)
                Dim idx As Long: idx = .ListCount - 1
                For ColIndex = 2 To 9
                    If ColIndex <> 6 Then
                        .List(idx, ColIndex - 1) = data(i, ColIndex)
                    Else
                        hresFormat = Format$(data(i, ColIndex), "#0.00")
                        hresFormat = Space(5 - Len(hresFormat)) & hresFormat
                        .List(idx, ColIndex - 1) = hresFormat
                    End If
                Next ColIndex
            End With
        Next i
        Set totaux = CalculerTotaux(data)
    End If
    
    'Mise à jour des totaux
    ufSaisieHeures.txtTotalHeures.Value = Format$(totaux(1), "#0.00")
    ufSaisieHeures.txtHresFact.Value = Format$(totaux(2), "#0.00")
    ufSaisieHeures.txtHresNF.Value = Format$(totaux(3), "#0.00")
    
    'Maintenant, on traite la semaine à partir de wshTEC_TDB_Data
    Dim totalHresFactSemaine As Currency
    Dim totalHresNonFactSemaine As Currency
    
    'Modifie les critères pour forcer une execution du AdvancedFilter dans wshTEC_TDB_Data
    Dim dateCharge As Date, dateLundi As Date, dateDimanche As Date
    dateCharge = ufSaisieHeures.txtDate.Value
    dateLundi = Fn_DateDuLundi(dateCharge)
    dateDimanche = dateLundi + 6
    Application.EnableEvents = False
    wshTEC_TDB_Data.Range("S7").Value = ufSaisieHeures.cmbProfessionnel.Value
    wshTEC_TDB_Data.Range("T7").Value = dateLundi
    Application.EnableEvents = True
    wshTEC_TDB_Data.Range("U7").Value = dateDimanche
    
    lastRow = wshTEC_TDB_Data.Cells(wshTEC_TDB_Data.Rows.count, "W").End(xlUp).Row
    If lastRow > 1 Then
        Set rngResult = wshTEC_TDB_Data.Range("W2:AD" & lastRow)
        totalHresFactSemaine = Application.WorksheetFunction.Sum(rngResult.Columns(7))
        totalHresNonFactSemaine = Application.WorksheetFunction.Sum(rngResult.Columns(8))
    End If

    ufSaisieHeures.txtHresFactSemaine.Value = Format$(totalHresFactSemaine, "#0.00")
    ufSaisieHeures.txtHresNFSemaine.Value = Format$(totalHresNonFactSemaine, "#0.00")
    
    ufSaisieHeures.Repaint
    
    DoEvents
    
    Application.ScreenUpdating = True

EndOfProcedure:

    ufSaisieHeures.txtClient.SetFocus
    
    'Libération et fin
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Set rngResult = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:RafraichirListBoxEtAddtionnerHeures", _
                        ufSaisieHeures.txtProfID.Value & " / " & ufSaisieHeures.txtDate.Value, startTime)
    Exit Sub
    
ErrorHandler:

    Call EnregistrerErreurs("modTEC_Saisie", "RafraichierListBoxEtAdditionnerHeures", "", Err.Number)
    MsgBox "Erreur : " & Err.description, vbCritical, "Erreur # APP-003"
    
    Resume EndOfProcedure
    
End Sub

Sub RafraichirTableauDeBordTEC()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:RafraichirTableauDeBordTEC", vbNullString, 0)

    Dim wsFrom As Worksheet: Set wsFrom = wsdTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = wsFrom.Cells(wsFrom.Rows.count, 1).End(xlUp).Row
    
    'Charger en mémoire toutes les données source
    Dim rawData As Variant
    rawData = wsFrom.Range("A3:N" & lastUsedRow).Value
    
    'Préparer le tableau des données à la sortie
    Dim arr() As Variant
    Dim numRows As Long: numRows = UBound(rawData, 1)
    ReDim arr(1 To numRows, 1 To 11)
    
    Dim i As Long
    For i = 1 To numRows
        arr(i, 1) = rawData(i, fTECTECID)
        arr(i, 2) = Format$(rawData(i, fTECProfID), "000")
        arr(i, 3) = rawData(i, fTECProf)
        arr(i, 4) = rawData(i, fTECDate)
        arr(i, 5) = rawData(i, fTECClientID)
        arr(i, 6) = rawData(i, fTECClientNom)
        arr(i, 7) = IIf(Fn_Is_Client_Facturable(rawData(i, fTECClientID)), "VRAI", "FAUX")
        arr(i, 8) = rawData(i, fTECHeures)
        arr(i, 9) = rawData(i, fTECEstFacturable)
        arr(i, 10) = rawData(i, fTECEstFacturee)
        arr(i, 11) = rawData(i, fTECEstDetruit)
    Next i
    
    'Mettre à jour la feuille TEC_TDB_Data
    Dim rngTo As Range
    Set rngTo = wshTEC_TDB_Data.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2))
    rngTo.Value = arr
    
    'Libérer la mémoire
    Set rngTo = Nothing
    Set wsFrom = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:RafraichirTableauDeBordTEC", vbNullString, startTime)

End Sub

Sub RafraichirTableauxCroisesTEC()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:RafraichirTableauxCroisesTEC", vbNullString, 0)

    Dim pt As PivotTable
    For Each pt In wshTEC_TDB.PivotTables
        pt.RefreshTable
    Next pt

    'Libérer la mémoire
    Set pt = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:RafraichirTableauxCroisesTEC", vbNullString, startTime)
    
End Sub

Sub ActiverButtonsVraiOuFaux(a As Boolean, u As Boolean, d As Boolean, c As Boolean)
                                  
    With ufSaisieHeures
        .shpAdd.Enabled = a
        .shpUpdate.Enabled = u
        .shpDelete.Enabled = d
        .shpClear.Enabled = c
    End With

End Sub

Sub MettreAJourPivotTables()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:MettreAJourPivotTables", vbNullString, 0)
    
    Dim ws As Worksheet: Set ws = wshStatsHeuresPivotTables
    Dim pt As PivotTable
    
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
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Saisie:MettreAJourPivotTables", vbNullString, startTime)
    
End Sub

Public Sub RemplirChampsRecordset(ByRef rs As Object, ByVal tecID As Long, ByVal dateValue As Date, _
                    ByVal timeStamp As Date, ByVal isUpdate As Boolean) '2025-10-31 @ 06:07
    
    With rs
        If Not isUpdate Then .Fields(fTECTECID - 1).Value = tecID
        .Fields(fTECProfID - 1).Value = ufSaisieHeures.txtProfID.Value
        .Fields(fTECProf - 1).Value = ufSaisieHeures.cmbProfessionnel.Value
        .Fields(fTECDate - 1).Value = dateValue
        .Fields(fTECClientID - 1).Value = ufSaisieHeures.txtClientID.Value
        .Fields(fTECClientNom - 1).Value = ufSaisieHeures.txtClient.Value
        .Fields(fTECDescription - 1).Value = Left$(ufSaisieHeures.txtActivite.Value, 255)
        .Fields(fTECHeures - 1).Value = Format$(ufSaisieHeures.txtHeures.Value, "#0.00")
        .Fields(fTECCommentaireNote - 1).Value = ufSaisieHeures.txtCommNote.Value
        .Fields(fTECEstFacturable - 1).Value = Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chkFacturable.Value)
        .Fields(fTECDateSaisie - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        .Fields(fTECEstFacturee - 1).Value = Fn_Convert_Value_Boolean_To_Text(False)
        .Fields(fTECDateFacturee - 1).Value = Null
        .Fields(fTECEstDetruit - 1).Value = Fn_Convert_Value_Boolean_To_Text(False)
        .Fields(fTECVersionApp - 1).Value = ThisWorkbook.Name
        .Fields(fTECNoFacture - 1).Value = vbNullString
    End With
    
End Sub

Public Function ConstruireLigneLog() As String '2025-10-31 @ 06:07

    ConstruireLigneLog = ufSaisieHeures.cmbProfessionnel.Value & " | " & _
                         Format$(ufSaisieHeures.txtDate.Value, "YYYY-MM-DD") & " | " & _
                         ufSaisieHeures.txtClientID.Value & " | " & _
                         ufSaisieHeures.txtClient.Value & " | " & _
                         ufSaisieHeures.txtActivite.Value & " | " & _
                         Format$(ufSaisieHeures.txtHeures.Value, "#0.00") & " | " & _
                         Fn_Convert_Value_Boolean_To_Text(ufSaisieHeures.chkFacturable.Value) & " | " & _
                         ufSaisieHeures.txtCommNote.Value
                         
End Function

Function CalculerTotaux(data As Variant) As Collection '2025-11-02 @ 09:58

    'Au cas où il n'y a rien...
    If IsEmpty(data) Or UBound(data, 1) = 0 Then
        Set CalculerTotaux = Nothing
        Exit Function
    End If
    
    Dim total As Currency, fact As Currency, nonFact As Currency
    Dim i As Long
    For i = 1 To UBound(data, 1)
        total = total + CCur(data(i, 6))
        If Fn_Is_Client_Facturable(data(i, 14)) Then
            fact = fact + CCur(data(i, 6))
        Else
            nonFact = nonFact + CCur(data(i, 6))
        End If
    Next i
    Dim result As New Collection
    result.Add total
    result.Add fact
    result.Add nonFact
    Set CalculerTotaux = result
    
End Function

