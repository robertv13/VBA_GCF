Attribute VB_Name = "modMain"
'@IgnoreModule ImplicitActiveWorkbookReference, HostSpecificExpression
'@Folder("Gestion_Clients")
Option Explicit

Public Const DATA_PATH As String = "\DataFiles"

Sub Show_Form()
    
    Call CM_Client_List_Import_All 'Toujours avoir la dernière version des clients

    ufClientMF.Show vbModeless

End Sub

Sub CM_Reset_UserForm()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Reset_UserForm", "", 0)
    
    Dim iRow As Long, lastUsedRow As Long
    iRow = Application.WorksheetFunction.CountA(Sheets("Donnees").Range("A:A"))
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
        .txtNomClientPlusNomClientSysteme.Value = ""
        
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
        .txtNomClientPlusNomClientSysteme.BackColor = vbWhite
        
        .txtRowNumber.Value = ""
        
        'Below code are associated with Search Feature
        If .txtSearch.Value <> "" Then
            Call CM_Build_Donnees_Recherche
            GoTo no_change
        End If
        Call CM_Add_SearchColumn
        'Worksheet - Donnees
        wshClients.AutoFilterMode = False
        'Worksheet - CM_Build_Donnees_Recherche
        wshSearchData.AutoFilterMode = False
        wshSearchData.Cells.Clear
        
        'ListBox parameters
        .lstDonnees.ColumnCount = 17
        .lstDonnees.ColumnHeads = True
        
        .lstDonnees.ColumnWidths = "200; 45; 150; 110; 110; 150; 130; 90; 95; 40; 55; 80; 100; 60; 105; 105; 350"
        
        '.RowSource
        Err.Clear
        On Error Resume Next
'        iRow = Application.WorksheetFunction.CountA(wshSearchData.Range("A:A"))
        If iRow > 1 Then
            .lstDonnees.RowSource = "Donnees!A2:R" & iRow
        End If
        If Err.Number = 380 Then
            MsgBox "Il y a un problème avec une commande de programmation." & _
                    vbNewLine & vbNewLine & _
                    "Valeur de iRow = " & iRow & _
                    vbNewLine & vbNewLine & _
                    "Sortir de l'application et réessayer à nouveau", _
                    vbCritical
        End If
        On Error GoTo 0

no_change:

    End With
    
    Call CM_Log_Activities("modMain:CM_Reset_UserForm", " Row=" & CStr(iRow), startTime)

End Sub

Sub CM_Ecrire_Client(action As String) '2025-06-27 @ 10:04

    Dim startTime As Double: startTime = Timer
    Call CM_Log_Activities("modMain:CM_Ecrire_Client", action, 0)

    'Lecture des données à une seule place
    Dim client As DonneesClient
    client = LireClientDepuisFormulaire(ufClientMF)

    'Écriture dans les deux sources
    Call CM_Update_External_GCF_Entree_BD(action, client)
    Call CM_Update_Locally_GCF_BD_Entree(action, client)

    'Log final avec identifiant client et durée
    Call CM_Log_Activities("modMain:CM_Ecrire_Client", action & " " & client.ClientID, startTime)

End Sub

Sub CM_Update_External_GCF_Entree_BD(action As String, client As DonneesClient) '2025-06-27 @ 10:05

    Dim startTime As Double: startTime = Timer
    Call CM_Log_Activities("modMain:Update_External_GCF_BD_Entree", action, 0)

    Application.ScreenUpdating = False

    Dim destinationFileName As String, destinationTab As String
    If Not Fn_Get_Windows_Username = "RobertMV" Then
        destinationFileName = "P:\Administration\APP\GCF\DataFiles\GCF_BD_Entrée.xlsx"
    Else
        destinationFileName = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
    End If
    destinationTab = "Clients$"

    'Vérifie si le fichier est ouvert
    If FichierEstOuvert(destinationFileName) Then
        MsgBox "Le classeur (GCF_BD_Entrée.xlsx) est actuellement utilisé." & vbNewLine & vbNewLine & _
               "Vous devez obligatoirement le fermer" & vbNewLine & vbNewLine & "avant de continuer.", _
               vbCritical, "Fichier est en cours d'utilisation"
        Exit Sub
    End If

    'Connexion ADO
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
              ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    If action = "NEW_RECORD" Then
        rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
        rs.AddNew
    Else
        rs.Open "SELECT * FROM [" & destinationTab & "] WHERE ClientID='" & client.ClientID & "'", conn, 2, 3
        If rs.EOF Then
            MsgBox "Le client '" & client.ClientID & "' n'a pas été ajouté au fichier!" & vbNewLine & vbNewLine & _
                   "Veuillez le saisir à nouveau", vbCritical, "ERREUR dans la mise à jour du fichier client"
            GoTo Nettoyage
        End If
    End If

    'Remplissage des champs
    With rs
        .Fields("ClientNom").Value = client.ClientNom
        .Fields("ClientID").Value = client.ClientID
        .Fields("NomClientSystème").Value = client.NomClientSystème
        .Fields("ContactFacturation").Value = client.ContactFacturation
        .Fields("TitreContactFacturation").Value = client.TitreContact
        .Fields("CourrielFacturation").Value = client.CourrielFacturation
        .Fields("Adresse1").Value = client.Adresse1
        .Fields("Adresse2").Value = client.Adresse2
        .Fields("Ville").Value = client.Ville
        .Fields("Province").Value = client.Province
        .Fields("CodePostal").Value = client.CodePostal
        .Fields("Pays").Value = client.Pays
        .Fields("RéféréPar").Value = client.ReferePar
        .Fields("FinAnnée").Value = client.FinAnnee
        .Fields("Comptable").Value = client.Comptable
        .Fields("NotaireAvocat").Value = client.NotaireAvocat
        .Fields("NomClientPlusNomClientSystème").Value = client.NomClientPlusNomClientSystème
        .Fields("TimeStamp").Value = client.TimeStamp
        .Update
    End With

Nettoyage:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close: Set rs = Nothing
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close: Set conn = Nothing
    On Error GoTo 0

    'Forcer Excel à enregistrer pour fiabilité
    Dim xlApp As Object, wb As Workbook
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False
    xlApp.Visible = False

    On Error Resume Next
    Set wb = xlApp.Workbooks.Open(destinationFileName, False, False)
    wb.Save
    wb.Close False
    xlApp.Quit
    Set wb = Nothing
    Set xlApp = Nothing
    On Error GoTo 0

    DoEvents

    'Vérification réelle de l’écriture sur disque
    Dim ddm As Date, jours As Long, heures As Long, minutes As Long, secondes As Long
    Call CM_Get_Date_Derniere_Modification(destinationFileName, ddm, jours, heures, minutes, secondes)

    If jours > 0 Or heures > 0 Or minutes > 0 Or secondes > 10 Then
        MsgBox "ATTENTION, le fichier MAÎTRE (GCF_Entrée.xlsx)" & vbNewLine & vbNewLine & _
               "n'a pas été modifié adéquatement sur disque..." & vbNewLine & vbNewLine & _
               "VEUILLEZ CONTACTER LE DÉVELOPPEUR SVP" & vbNewLine & vbNewLine & _
               "Code: (" & jours & "." & heures & "." & minutes & "." & secondes & ")", vbCritical, _
               "Le fichier n'est pas à jour sur disque"
    End If

    Application.ScreenUpdating = True

    Call CM_Log_Activities("modMain:Update_External_GCF_BD_Entree", action & " " & client.ClientID, startTime)

End Sub

Sub CM_Update_Locally_GCF_BD_Entree(action As String, client As DonneesClient) '2025-06-27 @ 10:05

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Update_Locally_GCF_BD_Entree", "", 0)

    Dim iRow As Long
    If ufClientMF.txtRowNumber.Value = "" Then
        iRow = Application.WorksheetFunction.CountA(Sheets("Donnees").Range("A:A")) + 1
    Else
        iRow = ufClientMF.txtRowNumber.Value
    End If

    With wshClients
        .Cells(iRow, 1) = client.ClientNom
        .Cells(iRow, 2) = client.ClientID
        .Cells(iRow, 3) = client.NomClientSystème
        .Cells(iRow, 4) = client.ContactFacturation
        .Cells(iRow, 5) = client.TitreContact
        .Cells(iRow, 6) = client.CourrielFacturation
        .Cells(iRow, 7) = client.Adresse1
        .Cells(iRow, 8) = client.Adresse2
        .Cells(iRow, 9) = client.Ville
        .Cells(iRow, 10) = client.Province
        .Cells(iRow, 11) = client.CodePostal
        .Cells(iRow, 12) = client.Pays
        .Cells(iRow, 13) = client.ReferePar
        .Cells(iRow, 14) = client.FinAnnee
        .Cells(iRow, 15) = client.Comptable
        .Cells(iRow, 16) = client.NotaireAvocat
        .Cells(iRow, 17) = client.NomClientPlusNomClientSystème
        .Cells(iRow, 18) = client.TimeStamp
    End With

    Call CM_Log_Activities("modMain:CM_Update_Locally_GCF_BD_Entree", action & " " & ufClientMF.txtCodeClient.Value, startTime)

End Sub

Sub CM_Add_SearchColumn()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Add_SearchColumn", "", 0)
    
    ufClientMF.EnableEvents = False

    With ufClientMF.cmbSearchColumn
        .Clear
        .AddItem "ClientID"
        .AddItem "ClientNom"
        .AddItem "NomClientSystème"
        .AddItem "ContactFacturation"
        .AddItem "TitreContactFacturation"
        .AddItem "CourrielFacturation"
        .AddItem "Adresse1"
        .AddItem "Adresse2"
        .AddItem "Ville"
        .AddItem "Province"
        .AddItem "CodePostal"
        .AddItem "Pays"
        .AddItem "RéféréPar"
        .AddItem "FinAnnée"
        .AddItem "Comptable"
        .AddItem "NotaireAvocat"
        .AddItem "NomClientPlusNomClientSystème"
        
        .Value = "ClientID"
    End With
    
    ufClientMF.EnableEvents = True
    
    ufClientMF.txtSearch.Value = ""
    ufClientMF.txtSearch.Enabled = True
'   ufClientMF.txtSearch.Enabled = False
    ufClientMF.cmdSearch.Enabled = False

    Call CM_Log_Activities("modMain:CM_Add_SearchColumn", "", startTime)

End Sub

Sub CM_Build_Donnees_Recherche()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Build_Donnees_Recherche", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim iColumn As Integer 'To hold the selected column number in Donnees sheet
    Dim iDonneesRow As Long 'To store the last non-blank row number available in Donnees sheet
    Dim iSearchRow As Long 'To hold the last non-blank row number available in SearchData sheet
    
    Dim sColumn As String 'To store the column selection
    Dim sValue As String 'To hold the search text value
    
    'Donnees sheet
    
    'DonneesRecherche sheet
    iDonneesRow = wshClients.Range("A" & Application.Rows.Count).End(xlUp).Row
    sColumn = ufClientMF.cmbSearchColumn.Value
    sValue = ufClientMF.txtSearch.Value
    iColumn = Application.WorksheetFunction.Match(sColumn, wshClients.Range("A1:R1"), 0)
    
    'Remove filter from Donnees worksheet
    If wshClients.FilterMode = True Then
        wshClients.AutoFilterMode = False
    End If

    'Apply filter on Donnees worksheet
    If ufClientMF.cmbSearchColumn.Value = "Code Client" Then
        wshClients.Range("A1:R" & iDonneesRow).AutoFilter Field:=iColumn, Criteria1:=sValue
    Else
        wshClients.Range("A1:R" & iDonneesRow).AutoFilter Field:=iColumn, Criteria1:="*" & sValue & "*"
    End If
    
    Dim searchRowsFound As Long
    searchRowsFound = Application.WorksheetFunction.Subtotal(3, wshClients.Range("A:A")) - 1 'Heading
    If searchRowsFound >= 1 Then
        'Code to remove the previous data from CM_Build_Donnees_Recherche worksheet
        wshSearchData.Cells.Clear
        wshClients.AutoFilter.Range.Copy wshSearchData.Range("A1")
        Application.CutCopyMode = False
        iSearchRow = wshSearchData.Range("A" & Application.Rows.Count).End(xlUp).Row
        ufClientMF.lstDonnees.ColumnCount = 17
        ufClientMF.lstDonnees.ColumnWidths = "200; 45; 150; 110; 110; 150; 130; 90; 95; 40; 55; 80; 100; 60; 105; 105; 350"
        If iSearchRow > 1 Then
            ufClientMF.lstDonnees.RowSource = "DonneesRecherche!A2:R" & iSearchRow
            ufClientMF.lblResultCount = "J'ai trouvé " & iSearchRow - 1 & " clients" '2024-08-24 @ 10:21
        End If
    Else
       MsgBox "Je n'ai trouvé AUCUN enregistrement avec ce critère."
    End If

    wshClients.AutoFilterMode = False
    Application.ScreenUpdating = True

    Call CM_Log_Activities("modMain:CM_Build_Donnees_Recherche", ufClientMF.cmbSearchColumn.Value & "=" & sValue & " " & searchRowsFound, startTime)

End Sub

Sub CM_Client_List_Import_All() 'Using ADODB - 2024-10-26 @ 12:05

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:CM_Client_List_Import_All", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshClients.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    If Not Fn_Get_Windows_Username = "RobertMV" Then
        sourceWorkbook = "P:\Administration\APP\GCF\DataFiles\GCF_BD_Entrée.xlsx"
    Else
        sourceWorkbook = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
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


