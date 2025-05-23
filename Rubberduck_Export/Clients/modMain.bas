Attribute VB_Name = "modMain"
'@IgnoreModule ImplicitActiveWorkbookReference, HostSpecificExpression
'@Folder("Gestion_Clients")
Option Explicit

Public Const DATA_PATH As String = "\DataFiles"

Sub Show_Form()
    
    Call CM_Client_List_Import_All 'Toujours avoir la derni�re version des clients

    ufClientMF.Show vbModeless

End Sub

Sub CM_Reset_UserForm()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Reset_UserForm", "", 0)
    
    Dim iRow As Long, lastUsedRow As Long
    iRow = Application.WorksheetFunction.CountA(Sheets("Donn�es").Range("A:A"))
'    iRow = [Counta(Donn�es!A:A)] 'Identifying the number of rows
    lastUsedRow = wshClients.Cells(wshClients.Rows.Count, "A").End(xlUp).Row
    
    With ufClientMF
        'Empty all fields
        .txtCodeClient.Value = ""
        .txtNomClient.Value = ""
        .txtNomClientSysteme = ""
        .txtContactFact.Value = ""
        .txtTitreContact.Value = ""
        .txtCourrielFact.Value = ""
        .txtAdresse1.Value = ""
        .txtAdresse2.Value = ""
        .txtVille.Value = ""
        .txtProvince.Value = ""
        .txtCodePostal.Value = ""
        .txtPays.Value = ""
        .txtReferePar.Value = ""
        .cmbFinAnnee.Value = ""
        .txtFinAnnee.Value = ""
        .txtComptable.Value = ""
        .txtNotaireAvocat.Value = ""
        .txtNomClientPlusNomClientSyst�me.Value = ""
        
        'Default Color for all fields
        .txtCodeClient.BackColor = vbWhite
        .txtNomClient.BackColor = vbWhite
        .txtNomClientSysteme.BackColor = vbWhite
        .txtContactFact.BackColor = vbWhite
        .txtTitreContact.BackColor = vbWhite
        .txtCourrielFact.BackColor = vbWhite
        .txtAdresse1.BackColor = vbWhite
        .txtAdresse2.BackColor = vbWhite
        .txtVille.BackColor = vbWhite
        .txtProvince.BackColor = vbWhite
        .txtCodePostal.BackColor = vbWhite
        .txtPays.BackColor = vbWhite
        .txtReferePar.BackColor = vbWhite
        .txtFinAnnee.BackColor = vbWhite
        .txtComptable.BackColor = vbWhite
        .txtNotaireAvocat.BackColor = vbWhite
        .txtNomClientPlusNomClientSyst�me.BackColor = vbWhite
        
        .txtRowNumber.Value = ""
        
        'Below code are associated with Search Feature
        If .txtSearch.Value <> "" Then
            Call CM_Build_Donn�es_Recherche
            GoTo no_change
        End If
        Call CM_Add_SearchColumn
        'Worksheet - Donn�es
        wshClients.AutoFilterMode = False
        'Worksheet - CM_Build_Donn�es_Recherche
        wshSearchData.AutoFilterMode = False
        wshSearchData.Cells.Clear
        
        'ListBox parameters
        .lstDonn�es.ColumnCount = 17
        .lstDonn�es.ColumnHeads = True
        
        .lstDonn�es.ColumnWidths = "200; 45; 150; 110; 110; 150; 130; 90; 95; 40; 55; 80; 100; 60; 105; 105; 350"
        
        '.RowSource
        On Error Resume Next
        If iRow > 1 Then
            .lstDonn�es.RowSource = "Donn�es!A2:R" & iRow
        End If
        If Err.Number = 380 Then
            MsgBox "Il y a un probl�me avec une commande de programmation." & _
                    vbNewLine & vbNewLine & _
                    "Valeur de iRow = " & iRow & _
                    vbNewLine & vbNewLine & _
                    "Sortir de l'application et r�essayer � nouveau" & _
                    vbCritical
        End If
        On Error GoTo 0

no_change:

    End With
    
    Call CM_Log_Activities("modMain:CM_Reset_UserForm", " Row=" & CStr(iRow), startTime)

End Sub

'CommentOut - 2025-03-13 @ 08:13
'Sub CM_Update_External_GCF_BD_Entr�e(action As String)
'
'    On Error GoTo ErrorHandler
'
'    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Update_External_GCF_BD_Entr�e", action, 0)
'
'    Application.ScreenUpdating = False
'    Application.Visible = False
'
'    'D�finir le nom du fichier en fonction de l'utilisateur (environnement)
'    Dim destinationFileName As String, destinationFileNamePath As String
'    If Not Fn_Get_Windows_Username = "Robert M. Vigneault" Then
'        destinationFileNamePath = "P:\Administration\APP\GCF\DataFiles"
'    Else
'        destinationFileNamePath = "C:\VBA\GC_FISCALIT�\DataFiles"
'    End If
'    destinationFileName = destinationFileNamePath & Application.PathSeparator & _
'                            "GCF_BD_Entr�e.xlsx"
'    Dim destinationTab As String: destinationTab = "Clients"
'
'    'Ouvrir le fichier GCF_BD_Entr�e.xlsx (could be PROD or DEV)
'    Dim wb As Workbook: Set wb = Workbooks.Open(destinationFileName, ReadOnly:=False)
'    Dim ws As Worksheet: Set ws = wb.Sheets(destinationTab)
'
'    Dim foundCell As Range
'
'    If action = "NEW_RECORD" Then
'        'Ajouter un nouvel enregistrement � la premi�re ligne vide (.Offset(1,0))
'        Set foundCell = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1, 0)
'        'L'offset Row est toujours � 0, et l'Offset Col varie de 0 @ 17
'        foundCell.Offset(0, 0).Value = ufClientMF.txtNomClient.Value
'        foundCell.Offset(0, 1).Value = ufClientMF.txtCodeClient.Value
'        foundCell.Offset(0, 2).Value = ufClientMF.txtNomClientSysteme.Value
'        foundCell.Offset(0, 3).Value = ufClientMF.txtContactFact.Value
'        foundCell.Offset(0, 4).Value = ufClientMF.txtTitreContact.Value
'        foundCell.Offset(0, 5).Value = ufClientMF.txtCourrielFact.Value
'        foundCell.Offset(0, 6).Value = ufClientMF.txtAdresse1.Value
'        foundCell.Offset(0, 7).Value = ufClientMF.txtAdresse2.Value
'        foundCell.Offset(0, 8).Value = ufClientMF.txtVille.Value
'        foundCell.Offset(0, 9).Value = ufClientMF.txtProvince.Value
'        foundCell.Offset(0, 10).Value = ufClientMF.txtCodePostal.Value
'        foundCell.Offset(0, 11).Value = ufClientMF.txtPays.Value
'        foundCell.Offset(0, 12).Value = ufClientMF.txtReferePar.Value
'        foundCell.Offset(0, 13).Value = ufClientMF.txtFinAnnee.Value
'        foundCell.Offset(0, 14).Value = ufClientMF.txtComptable.Value
'        foundCell.Offset(0, 15).Value = ufClientMF.txtNotaireAvocat.Value
'        foundCell.Offset(0, 16).Value = ufClientMF.txtNomClientPlusNomClientSyst�me.Value
'        foundCell.Offset(0, 17).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    Else
'        'Rechercher le client existant par son ID, dans la 2�me colonne
'        Set foundCell = ws.Range("B:B").Find(ufClientMF.txtCodeClient.Value, LookIn:=xlValues, LookAt:=xlWhole)
'        If Not foundCell Is Nothing Then
'            'Modifier les champs de l'enregistrement existant
'            foundCell.Offset(0, -1).Value = ufClientMF.txtNomClient.Value
'            foundCell.Offset(0, 0).Value = ufClientMF.txtCodeClient.Value
'            foundCell.Offset(0, 1).Value = ufClientMF.txtNomClientSysteme.Value
'            foundCell.Offset(0, 2).Value = ufClientMF.txtContactFact.Value
'            foundCell.Offset(0, 3).Value = ufClientMF.txtTitreContact.Value
'            foundCell.Offset(0, 4).Value = ufClientMF.txtCourrielFact.Value
'            foundCell.Offset(0, 5).Value = ufClientMF.txtAdresse1.Value
'            foundCell.Offset(0, 6).Value = ufClientMF.txtAdresse2.Value
'            foundCell.Offset(0, 7).Value = ufClientMF.txtVille.Value
'            foundCell.Offset(0, 8).Value = ufClientMF.txtProvince.Value
'            foundCell.Offset(0, 9).Value = ufClientMF.txtCodePostal.Value
'            foundCell.Offset(0, 10).Value = ufClientMF.txtPays.Value
'            foundCell.Offset(0, 11).Value = ufClientMF.txtReferePar.Value
'            foundCell.Offset(0, 12).Value = ufClientMF.txtFinAnnee.Value
'            foundCell.Offset(0, 13).Value = ufClientMF.txtComptable.Value
'            foundCell.Offset(0, 14).Value = ufClientMF.txtNotaireAvocat.Value
'            foundCell.Offset(0, 15).Value = ufClientMF.txtNomClientPlusNomClientSyst�me.Value
'            foundCell.Offset(0, 16).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'        Else
'            MsgBox "#MB:101 - Le client '" & ufClientMF.txtCodeClient & "' n'a pas �t� trouv� dans le fichier!", vbCritical
'        End If
'    End If
'
'    'Ferme ET sauvegarde le fichier Excel
'    wb.Close SaveChanges:=True
'
'CleanUp:
'    'Nettoyer les ressources
'    Set foundCell = Nothing
'    Set ws = Nothing
'    Set wb = Nothing
'
'    'Is the file really modified on disk ?
'    Call CM_Verify_DDM(destinationFileName)
'
'    'Restauration des param�tres Excel
'    Application.Visible = True
'    Application.ScreenUpdating = True
'
'    Call CM_Log_Activities("modMain:CM_Update_External_GCF_BD_Entr�e", action & " " & ufClientMF.txtCodeClient.Value, startTime)
'
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "Erreur: " & Err.Description, vbCritical
'    Resume CleanUp
'
'End Sub
'
Sub CM_Update_External_GCF_Entr�e_BD(action As String) 'Update/Write Client record to Clients' Master File

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:Update_External_GCF_BD_Entree", action, 0)

    Application.ScreenUpdating = False

    Dim destinationFileName As String, destinationTab As String
    If Not Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        destinationFileName = "P:\Administration\APP\GCF\DataFiles\GCF_BD_Entr�e.xlsx"
    Else
        destinationFileName = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    End If
    destinationTab = "Clients$"
    
    'Test pour s'assurer que le classeur n'est pas ouvert (cr�ation de probl�me avec ADO)
    If FichierEstOuvert(destinationFileName) Then
        MsgBox "Le classeur (GCF_BD_Entr�e.xlsx) est actuellement utilis�." & _
               vbNewLine & vbNewLine & "Vous devez obligatoirement le fermer" & _
               vbNewLine & vbNewLine & "avant de continuer.", vbCritical, "Fichier est en cours d'utilisation"
        Exit Sub
    End If
    
    Dim timeStamp As Date
    timeStamp = Format$(Now(), "yyyy-mm-dd hh:mm:ss")

    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    
    If action = "NEW_RECORD" Then
        'Open an empty recordset
        rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
        'Add fields to the recordset before updating it
        rs.AddNew
            rs.Fields("ClientNom").Value = ufClientMF.txtNomClient.Value
            rs.Fields("ClientID").Value = ufClientMF.txtCodeClient.Value
            rs.Fields("NomClientSyst�me").Value = ufClientMF.txtNomClientSysteme
            rs.Fields("ContactFacturation").Value = ufClientMF.txtContactFact.Value
            rs.Fields("TitreContactFacturation").Value = ufClientMF.txtTitreContact.Value
            rs.Fields("CourrielFacturation").Value = ufClientMF.txtCourrielFact.Value
            rs.Fields("Adresse1").Value = ufClientMF.txtAdresse1.Value
            rs.Fields("Adresse2").Value = ufClientMF.txtAdresse2.Value
            rs.Fields("Ville").Value = ufClientMF.txtVille.Value
            rs.Fields("Province").Value = ufClientMF.txtProvince.Value
            rs.Fields("CodePostal").Value = ufClientMF.txtCodePostal.Value
            rs.Fields("Pays").Value = ufClientMF.txtPays.Value
            rs.Fields("R�f�r�Par").Value = ufClientMF.txtReferePar.Value
            rs.Fields("FinAnn�e").Value = ufClientMF.txtFinAnnee.Value
            rs.Fields("Comptable").Value = ufClientMF.txtComptable.Value
            rs.Fields("NotaireAvocat").Value = ufClientMF.txtNotaireAvocat.Value
            rs.Fields("NomClientPlusNomClientSyst�me").Value = ufClientMF.txtNomClientPlusNomClientSyst�me.Value
            rs.Fields("TimeStamp").Value = timeStamp
            Debug.Print "Tous les champs ont �t� assign�s"
        rs.Update
        
    Else 'Update an existing record
        'Open the recordset for the existing client
        rs.Open "SELECT * FROM [" & destinationTab & "] WHERE ClientID='" & ufClientMF.txtCodeClient & "'", conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
            rs.Fields("ClientNom").Value = ufClientMF.txtNomClient.Value
            rs.Fields("ClientID").Value = ufClientMF.txtCodeClient.Value
            rs.Fields("NomClientSyst�me").Value = ufClientMF.txtNomClientSysteme
            rs.Fields("ContactFacturation").Value = ufClientMF.txtContactFact.Value
            rs.Fields("TitreContactFacturation").Value = ufClientMF.txtTitreContact.Value
            rs.Fields("CourrielFacturation").Value = ufClientMF.txtCourrielFact.Value
            rs.Fields("Adresse1").Value = ufClientMF.txtAdresse1.Value
            rs.Fields("Adresse2").Value = ufClientMF.txtAdresse2.Value
            rs.Fields("Ville").Value = ufClientMF.txtVille.Value
            rs.Fields("Province").Value = ufClientMF.txtProvince.Value
            rs.Fields("CodePostal").Value = ufClientMF.txtCodePostal.Value
            rs.Fields("Pays").Value = ufClientMF.txtPays.Value
            rs.Fields("R�f�r�Par").Value = ufClientMF.txtReferePar.Value
            rs.Fields("FinAnn�e").Value = ufClientMF.txtFinAnnee.Value
            rs.Fields("Comptable").Value = ufClientMF.txtComptable.Value
            rs.Fields("NotaireAvocat").Value = ufClientMF.txtNotaireAvocat.Value
            rs.Fields("NomClientPlusNomClientSyst�me").Value = ufClientMF.txtNomClientPlusNomClientSyst�me.Value
            rs.Fields("TimeStamp").Value = timeStamp
            rs.Update
            
'            If Err.Number <> 0 Then
'            MsgBox "Erreur lors de la mise � jour: " & Err.Description, vbCritical, "ERREUR dans la mise � jour du fichier client"
'            End If
'
'            Call CM_Log_Activities("modMain:Update_External_GCF_BD_Entree", action & " '" & ufClientMF.txtCodeClient.Value & "' was here", -1)
        Else
            'Handle the case where the specified ID is not found
            MsgBox "Le client '" & ufClientMF.txtCodeClient & "' n'a pas �t� ajout� au fichier!" & _
                    vbNewLine & vbNewLine & "Veuillez le saisir � nouveau", vbCritical, "ERREUR dans la mise � jour du fichier client"
        End If
    End If

    'Ferme proprement la connexion et le recordset
    If Not rs Is Nothing Then
        If rs.State = 1 Then
            rs.Close
            Set rs = Nothing
        End If
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then
            conn.Close
            Set conn = Nothing
        End If
    End If
    
    'Proposition chatGPT pour forcer l'�criture par EXCEL - 2025-05-02 @ 14:06
    'Ouvrir le fichier temporairement via Excel pour forcer l'enregistrement
    Dim xlApp As Object, wb As Workbook
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False
    xlApp.Visible = False

    On Error GoTo CleanupFile
    Set wb = xlApp.Workbooks.Open(destinationFileName, False, False)
    wb.Save
    wb.Close False

CleanupFile:
    xlApp.Quit
    Set wb = Nothing
    Set xlApp = Nothing
    
    DoEvents

'    Application.Wait Now + TimeValue("00:00:01") 'Attente de 1 seconde

    DoEvents
    
    'Additional verification - Actual file MUST have been modified (GCF_Entr�e.xlsx)
    Dim ddm As Date, jours As Long, heures As Long, minutes As Long, secondes As Long
    Call CM_Get_Date_Derniere_Modification(destinationFileName, _
                                                ddm, jours, heures, minutes, secondes)
    'Record to the log the difference between NOW and the date of last modifcation
    Call CM_Log_Activities("modMain:Update_External_GCF_BD_Entree", "DDM (" & ddm & " : " & jours & "." & heures & "." & minutes & "." & secondes & ")", -1)
    If jours > 0 Or heures > 0 Or minutes > 0 Or secondes > 10 Then
        MsgBox "ATTENTION, le fichier MA�TRE (GCF_Entr�e.xlsx)" & vbNewLine & vbNewLine & _
               "n'a pas �t� modifi� ad�quatement sur disque..." & vbNewLine & vbNewLine & _
               "VEUILLEZ CONTACTER LE D�VELOPPEUR SVP" & vbNewLine & vbNewLine & _
               "Code: (" & jours & "." & heures & "." & minutes & "." & secondes & ")", vbCritical, _
               "Le fichier n'est pas � jour sur disque"
    End If

    Application.ScreenUpdating = True

    Call CM_Log_Activities("modMain:Update_External_GCF_BD_Entree", action & " " & ufClientMF.txtCodeClient.Value, startTime)

End Sub

Sub CM_Update_Locally_GCF_BD_Entr�e(action As String)

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Update_Locally_GCF_BD_Entr�e", "", 0)
    
    Dim iRow As Long
    If ufClientMF.txtRowNumber.Value = "" Then
        iRow = Application.WorksheetFunction.CountA(Sheets("Donn�es").Range("A:A")) + 1
    Else
        iRow = ufClientMF.txtRowNumber.Value
    End If
    
    With wshClients
        .Cells(iRow, 1) = ufClientMF.txtNomClient.Value
        .Cells(iRow, 2) = ufClientMF.txtCodeClient.Value
        .Cells(iRow, 3) = ufClientMF.txtNomClientSysteme.Value
        .Cells(iRow, 4) = ufClientMF.txtContactFact.Value
        .Cells(iRow, 5) = ufClientMF.txtTitreContact.Value
        .Cells(iRow, 6) = ufClientMF.txtCourrielFact.Value
        .Cells(iRow, 7) = ufClientMF.txtAdresse1.Value
        .Cells(iRow, 8) = ufClientMF.txtAdresse2.Value
        .Cells(iRow, 9) = ufClientMF.txtVille.Value
        .Cells(iRow, 10) = ufClientMF.txtProvince.Value
        .Cells(iRow, 11) = ufClientMF.txtCodePostal.Value
        .Cells(iRow, 12) = ufClientMF.txtPays.Value
        .Cells(iRow, 13) = ufClientMF.txtReferePar.Value
        .Cells(iRow, 14) = ufClientMF.txtFinAnnee.Value
        .Cells(iRow, 15) = ufClientMF.txtComptable.Value
        .Cells(iRow, 16) = ufClientMF.txtNotaireAvocat.Value
        .Cells(iRow, 17) = ufClientMF.txtNomClientPlusNomClientSyst�me
        .Cells(iRow, 18) = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    End With

    Call CM_Log_Activities("modMain:CM_Update_Locally_GCF_BD_Entr�e", action & " " & ufClientMF.txtCodeClient.Value, startTime)

End Sub

Sub CM_Add_SearchColumn()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Add_SearchColumn", "", 0)
    
    ufClientMF.EnableEvents = False

    With ufClientMF.cmbSearchColumn
        .Clear
        .AddItem "ClientID"
        .AddItem "ClientNom"
        .AddItem "NomCLientSyst�me"
        .AddItem "ContactFacturation"
        .AddItem "TitreContactFacturation"
        .AddItem "CourrielFacturation"
        .AddItem "Adresse1"
        .AddItem "Adresse2"
        .AddItem "Ville"
        .AddItem "Province"
        .AddItem "CodePostal"
        .AddItem "Pays"
        .AddItem "R�f�r�Par"
        .AddItem "FinAnn�e"
        .AddItem "Comptable"
        .AddItem "NotaireAvocat"
        .AddItem "NomClientPlusNomClientSyst�me"
        
        .Value = "ClientID"
    End With
    
    ufClientMF.EnableEvents = True
    
    ufClientMF.txtSearch.Value = ""
    ufClientMF.txtSearch.Enabled = True
'    ufClientMF.txtSearch.Enabled = False
    ufClientMF.cmdSearch.Enabled = False

    Call CM_Log_Activities("modMain:CM_Add_SearchColumn", "", startTime)

End Sub

Sub CM_Build_Donn�es_Recherche()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Build_Donn�es_Recherche", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim iColumn As Integer 'To hold the selected column number in Donn�es sheet
    Dim iDonn�esRow As Long 'To store the last non-blank row number available in Donn�es sheet
    Dim iSearchRow As Long 'To hold the last non-blank row number available in SearchData sheet
    
    Dim sColumn As String 'To store the column selection
    Dim sValue As String 'To hold the search text value
    
     'Donn�es sheet
    
     'Donn�esRecherche sheet
    
    
    iDonn�esRow = wshClients.Range("A" & Application.Rows.Count).End(xlUp).Row
    sColumn = ufClientMF.cmbSearchColumn.Value
    sValue = ufClientMF.txtSearch.Value
    iColumn = Application.WorksheetFunction.Match(sColumn, wshClients.Range("A1:R1"), 0)
    
    'Remove filter from Donn�es worksheet
    If wshClients.FilterMode = True Then
        wshClients.AutoFilterMode = False
    End If

    'Apply filter on Donn�es worksheet
    If ufClientMF.cmbSearchColumn.Value = "Code Client" Then
        wshClients.Range("A1:R" & iDonn�esRow).AutoFilter Field:=iColumn, Criteria1:=sValue
    Else
        wshClients.Range("A1:R" & iDonn�esRow).AutoFilter Field:=iColumn, Criteria1:="*" & sValue & "*"
    End If
    
    Dim searchRowsFound As Long
    searchRowsFound = Application.WorksheetFunction.Subtotal(3, wshClients.Range("A:A")) - 1 'Heading
    If searchRowsFound >= 1 Then
        'Code to remove the previous data from CM_Build_Donn�es_Recherche worksheet
        wshSearchData.Cells.Clear
        wshClients.AutoFilter.Range.Copy wshSearchData.Range("A1")
        Application.CutCopyMode = False
        iSearchRow = wshSearchData.Range("A" & Application.Rows.Count).End(xlUp).Row
        ufClientMF.lstDonn�es.ColumnCount = 17
        ufClientMF.lstDonn�es.ColumnWidths = "200; 45; 150; 110; 110; 150; 130; 90; 95; 40; 55; 80; 100; 60; 105; 105; 350"
        If iSearchRow > 1 Then
            ufClientMF.lstDonn�es.RowSource = "Donn�esRecherche!A2:R" & iSearchRow
            ufClientMF.lblResultCount = "J'ai trouv� " & iSearchRow - 1 & " clients" '2024-08-24 @ 10:21
        End If
    Else
       MsgBox "Je n'ai trouv� AUCUN enregistrement avec ce crit�re."
    End If

    wshClients.AutoFilterMode = False
    Application.ScreenUpdating = True

    Call CM_Log_Activities("modMain:CM_Build_Donn�es_Recherche", ufClientMF.cmbSearchColumn.Value & "=" & sValue & " " & searchRowsFound, startTime)

End Sub

Sub CM_Client_List_Import_All() 'Using ADODB - 2024-10-26 @ 12:05

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Client_List_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshClients.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_BD_Entr�e.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    If Not Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        sourceWorkbook = "P:\Administration\APP\GCF\DataFiles\GCF_BD_Entr�e.xlsx"
    Else
        sourceWorkbook = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    End If
    sourceTab = "Clients$"
    
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    
    recSet.ActiveConnection = connStr
    recSet.Source = "SELECT * FROM [" & sourceTab & "]"
    recSet.Open
    
    'Copy to wshBD_Clients workbook
    wshClients.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet - 2024-07-20 @ 18:31
    Dim rng As Range: Set rng = wshClients.Range("A1").CurrentRegion
    
    Call CM_Apply_Worksheet_Format(wshClients, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    
    Call CM_Log_Activities("modMain:CM_Client_List_Import_All", "", startTime)

End Sub

Sub CM_Apply_Worksheet_Format(ws As Worksheet, rng As Range, headerRow As Long)

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Apply_Worksheet_Format", "", 0)
    
    'Common stuff to all worksheets
    rng.EntireColumn.AutoFit 'Autofit all columns
    
    'Conditional Formatting (many steps)
    '1) Remove existing conditional formatting
        rng.Cells.FormatConditions.Delete 'Remove the worksheet conditional formatting
    
    '2) Define the usedRange to data only (exclude header row(s))
        Dim numRows As Long
        numRows = rng.CurrentRegion.Rows.Count - headerRow
        Dim usedRange As Range
        If numRows > 0 Then
            On Error Resume Next
            Set usedRange = rng.Offset(headerRow, 0).Resize(numRows, rng.Columns.Count)
            On Error GoTo 0
        End If
    
    '3) Add the standard conditional formatting
        If Not usedRange Is Nothing Then
            usedRange.FormatConditions.Add Type:=xlExpression, _
                Formula1:="=ET($A2<>"""";mod(LIGNE();2)=1)"
            usedRange.FormatConditions(usedRange.FormatConditions.Count).SetFirstPriority
            With usedRange.FormatConditions(1).Font
                .Strikethrough = False
                .TintAndShade = 0
            End With
            With usedRange.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
            End With
            usedRange.FormatConditions(1).StopIfTrue = False
        End If
    
    Call CM_Log_Activities("modMain:CM_Apply_Worksheet_Format", CStr(numRows), startTime)

End Sub


