Attribute VB_Name = "modMain"
Option Explicit

Dim InitialValues As Collection

Sub Show_Form()
    
    Call Client_List_Import_All 'Toujours avoir la dernière version des clients

    frmForm.Show vbModeless

End Sub

Sub Reset()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Reset", 0)
    
    Dim iRow As Long
    iRow = [Counta(Données!A:A)] 'Identifying the number of rows
    
    With frmForm
        .txtCodeClient.Value = ""
        .txtNomClient.Value = ""
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
        
        'Default Color
        .txtCodeClient.BackColor = vbWhite
        .txtNomClient.BackColor = vbWhite
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
        
        .txtRowNumber.Value = ""
        
        'Below code are associated with Search Feature - Part 3
        Call Add_SearchColumn
        ThisWorkbook.Sheets("Données").AutoFilterMode = False
        ThisWorkbook.Sheets("DonnéesRecherche").AutoFilterMode = False
        ThisWorkbook.Sheets("DonnéesRecherche").Cells.Clear
        
        .lstDonnées.ColumnCount = 15
        .lstDonnées.ColumnHeads = True
        
        .lstDonnées.ColumnWidths = "200; 45; 110; 110; 150; 130; 90; 95; 40; 55; 80; 100; 70; 105; 105"
        
        'Try .RowSource...
        On Error Resume Next
        If iRow > 1 Then
            .lstDonnées.RowSource = "Données!A2:O" & iRow
'        Else
'            .lstDonnées.RowSource = "Données!A2:O2"
        End If
        If Err.Number = 380 Then
            MsgBox "Il y a un problème avec une commande de programmation." & _
                    vbNewLine & vbNewLine & _
                    "Valeur de iRow = " & iRow & _
                    vbNewLine & vbNewLine & _
                    "Sortir de l'application et réessayer à nouveau" & _
                    vbCritical
        End If
        On Error GoTo 0
            
    End With
    
    Call Log_Record("modMain:Reset", startTime)

End Sub

Sub Submit_GCF_BD_Entrée_Clients(action As String) 'Update/Write Client record to Clients' Master File

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Submit_GCF_BD_Entrée_Clients", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    If Not Environ("userName") = "Robert M. Vigneault" Then
        destinationFileName = "P:\Administration\APP\GCF\DataFiles\GCF_BD_Entrée.xlsx"
    Else
        destinationFileName = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
    End If
    destinationTab = "Clients"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    If action = "NEW_RECORD" Then
        
        'Open an empty recordset
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
        
        'Add fields to the recordset before updating it
        rs.AddNew
        rs.Fields("ClientNom").Value = frmForm.txtNomClient.Value
        rs.Fields("Client_ID").Value = frmForm.txtCodeClient.Value
        rs.Fields("ContactFacturation").Value = frmForm.txtContactFact.Value
        rs.Fields("TitreContactFacturation").Value = frmForm.txtTitreContact.Value
        rs.Fields("CourrielFacturation").Value = frmForm.txtCourrielFact.Value
        rs.Fields("Adresse_1").Value = frmForm.txtAdresse1.Value
        rs.Fields("Adresse_2").Value = frmForm.txtAdresse2.Value
        rs.Fields("Ville").Value = frmForm.txtVille.Value
        rs.Fields("Province").Value = frmForm.txtProvince.Value
        rs.Fields("CodePostal").Value = frmForm.txtCodePostal.Value
        rs.Fields("Pays").Value = frmForm.txtPays.Value
        rs.Fields("Référé par").Value = frmForm.txtReferePar.Value
        rs.Fields("Fin d'année").Value = frmForm.txtFinAnnee.Value
        rs.Fields("Comptable").Value = frmForm.txtComptable.Value
        rs.Fields("Notaire/Avocat").Value = frmForm.txtNotaireAvocat.Value
        
    Else 'Update an existing record
        'Open the recordset for the existing client
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE Client_ID='" & frmForm.txtCodeClient & "'", conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
            rs.Fields("ClientNom").Value = frmForm.txtNomClient.Value
            rs.Fields("Client_ID").Value = frmForm.txtCodeClient.Value
            rs.Fields("ContactFacturation").Value = frmForm.txtContactFact.Value
            rs.Fields("TitreContactFacturation").Value = frmForm.txtTitreContact.Value
            rs.Fields("CourrielFacturation").Value = frmForm.txtCourrielFact.Value
            rs.Fields("Adresse_1").Value = frmForm.txtAdresse1.Value
            rs.Fields("Adresse_2").Value = frmForm.txtAdresse2.Value
            rs.Fields("Ville").Value = frmForm.txtVille.Value
            rs.Fields("Province").Value = frmForm.txtProvince.Value
            rs.Fields("CodePostal").Value = frmForm.txtCodePostal.Value
            rs.Fields("Pays").Value = frmForm.txtPays.Value
            rs.Fields("Référé par").Value = frmForm.txtReferePar.Value
            rs.Fields("Fin d'année").Value = frmForm.txtFinAnnee.Value
            rs.Fields("Comptable").Value = frmForm.txtComptable.Value
            rs.Fields("Notaire/Avocat").Value = frmForm.txtNotaireAvocat.Value
        Else
            'Handle the case where the specified ID is not found
            MsgBox "Le client '" & frmForm.txtCodeClient & "' n'a pas été ajouté au fichier!" & _
                    vbNewLine & vbNewLine & "Veuillez le saisir à nouveau", vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    End If
    'Update the recordset (create the record)
    rs.Update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modMain:Submit_GCF_BD_Entrée_Clients", startTime)

End Sub

Sub Submit_Locally(action As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Submit_Locally", 0)
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Données")
    
    Dim iRow As Long
    If frmForm.txtRowNumber.Value = "" Then
        iRow = [Counta(Données!A:A)] + 1
    Else
        iRow = frmForm.txtRowNumber.Value
    End If
    
    With sh
        .Cells(iRow, 1) = frmForm.txtNomClient.Value
        .Cells(iRow, 2) = frmForm.txtCodeClient.Value
        .Cells(iRow, 3) = frmForm.txtContactFact.Value
        .Cells(iRow, 4) = frmForm.txtTitreContact.Value
        .Cells(iRow, 5) = frmForm.txtCourrielFact.Value
        .Cells(iRow, 6) = frmForm.txtAdresse1.Value
        .Cells(iRow, 7) = frmForm.txtAdresse2.Value
        .Cells(iRow, 8) = frmForm.txtVille.Value
        .Cells(iRow, 9) = frmForm.txtProvince.Value
        .Cells(iRow, 10) = frmForm.txtCodePostal.Value
        .Cells(iRow, 11) = frmForm.txtPays.Value
        .Cells(iRow, 12) = frmForm.txtReferePar.Value
        .Cells(iRow, 13) = frmForm.txtFinAnnee.Value
        .Cells(iRow, 14) = frmForm.txtComptable.Value
        .Cells(iRow, 15) = frmForm.txtNotaireAvocat.Value
'        .Cells(iRow, 9) = [Text(Now(), "DD-MM-YYYY HH:MM:SS")]
    End With

    Call Log_Record("modMain:Submit_Locally", startTime)

End Sub

Function Selected_List() As Long

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Selected_List", 0)
    
    Selected_List = 0
    
    Dim i As Long
    For i = 0 To frmForm.lstDonnées.ListCount - 1
        If frmForm.lstDonnées.Selected(i) = True Then
            Selected_List = i + 1
            frmForm.cmdEdit.Enabled = True
            Exit For
        End If
        frmForm.cmdEdit.Enabled = False
    Next i

    Call Log_Record("modMain:Selected_List", startTime)

End Function

Sub Add_SearchColumn()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Add_SearchColumn", 0)
    
    frmForm.EnableEvents = False

    With frmForm.cmbSearchColumn
        .Clear
        .AddItem "ClientNom"
        .AddItem "Client_ID"
        .AddItem "ContactFacturation"
        .AddItem "TitreContactFacturation"
        .AddItem "CourrielFacturation"
        .AddItem "Adresse_1"
        .AddItem "Adresse_2"
        .AddItem "Ville"
        .AddItem "Province"
        .AddItem "CodePostal"
        .AddItem "Pays"
        .AddItem "Référé par"
        .AddItem "Fin d'année"
        .AddItem "Comptable"
        .AddItem "Notaire/Avocat"
        
        .Value = "Client_ID"
        
    End With
    
    frmForm.EnableEvents = True
    
    frmForm.txtSearch.Value = ""
    frmForm.txtSearch.Enabled = True
'    frmForm.txtSearch.Enabled = False
    frmForm.cmdSearch.Enabled = False

    Call Log_Record("modMain:Add_SearchColumn", startTime)

End Sub

Sub DonnéesRecherche()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:DonnéesRecherche", 0)
    
    Application.ScreenUpdating = False
    
    Dim iColumn As Integer 'To hold the selected column number in Données sheet
    Dim iDonnéesRow As Long 'To store the last non-blank row number available in Données sheet
    Dim iSearchRow As Long 'To hold the last non-blank row number available in SearachData sheet
    
    Dim sColumn As String 'To store the column selection
    Dim sValue As String 'To hold the search text value
    
    Dim wshDonnées As Worksheet 'Données sheet
    Set wshDonnées = ThisWorkbook.Sheets("Données")
    Dim wshSearchData As Worksheet 'DonnéesRecherche sheet
    Set wshSearchData = ThisWorkbook.Sheets("DonnéesRecherche")
    
    iDonnéesRow = ThisWorkbook.Sheets("Données").Range("A" & Application.Rows.Count).End(xlUp).Row
    sColumn = frmForm.cmbSearchColumn.Value
    sValue = frmForm.txtSearch.Value
    iColumn = Application.WorksheetFunction.Match(sColumn, wshDonnées.Range("A1:O1"), 0)
    
    'Remove filter from Données worksheet
    If wshDonnées.FilterMode = True Then
        wshDonnées.AutoFilterMode = False
    End If

    'Apply filter on Données worksheet
    If frmForm.cmbSearchColumn.Value = "Code Client" Then
        wshDonnées.Range("A1:O" & iDonnéesRow).AutoFilter Field:=iColumn, Criteria1:=sValue
    Else
        wshDonnées.Range("A1:O" & iDonnéesRow).AutoFilter Field:=iColumn, Criteria1:="*" & sValue & "*"
    End If
    
    Dim searchRowsFound As Long
    searchRowsFound = Application.WorksheetFunction.Subtotal(3, wshDonnées.Range("A:A")) - 1 'Heading
    If searchRowsFound >= 1 Then
        'Code to remove the previous data from DonnéesRecherche worksheet
        wshSearchData.Cells.Clear
        wshDonnées.AutoFilter.Range.Copy wshSearchData.Range("A1")
        Application.CutCopyMode = False
        iSearchRow = wshSearchData.Range("A" & Application.Rows.Count).End(xlUp).Row
        frmForm.lstDonnées.ColumnCount = 15
        frmForm.lstDonnées.ColumnWidths = "200; 45; 110; 110; 150; 130; 90; 95; 40; 55; 80; 100; 70; 105; 105"
        If iSearchRow > 1 Then
            frmForm.lstDonnées.RowSource = "DonnéesRecherche!A2:O" & iSearchRow
'            MsgBox "J'ai trouvé " & searchRowsFound & " enregistrements."
        End If
    Else
       MsgBox "Je n'ai trouvé AUCUN enregistrement."
    End If

    wshDonnées.AutoFilterMode = False
    Application.ScreenUpdating = True

    Call Log_Record("modMain:DonnéesRecherche", startTime)

End Sub

Function ValidateEntries() As Boolean

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:ValidateEntries", 0)
    
    ValidateEntries = True
    
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Données")

    Dim iCodeClient As Variant
    iCodeClient = frmForm.txtCodeClient.Value
    
    With frmForm
        'Default Color
        .txtCodeClient.BackColor = vbWhite
        .txtNomClient.BackColor = vbWhite
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
        
        'Valeur OBLIGATOIRE
        If Trim(.txtCodeClient.Value) = "" Then
            MsgBox "SVP, saisir un code de client.", vbOKOnly + vbInformation, "Code de client"
            ValidateEntries = False
            .txtCodeClient.BackColor = vbRed
            .txtCodeClient.Enabled = True
            .txtCodeClient.SetFocus
            Exit Function
        End If
    
        'Valeur OBLIGATOIRE
        If Trim(.txtNomClient.Value) = "" Then
            MsgBox "SVP, saisir le nom du client.", vbOKOnly + vbInformation, "Nom de client"
            ValidateEntries = False
            .txtNomClient.BackColor = vbRed
            .txtNomClient.SetFocus
            Exit Function
        End If
        
    End With

    Call Log_Record("modMain:ValidateEntries", startTime)

End Function

Sub Client_List_Import_All() 'Using ADODB - 2024-08-07 @ 11:55

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Client_List_Import_All", 0)
    
    Application.StatusBar = "J'importe la liste des clients"
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wshClients.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    If Not Environ("userName") = "Robert M. Vigneault" Then
        sourceWorkbook = "P:\Administration\APP\GCF\DataFiles\GCF_BD_Entrée.xlsx"
    Else
        sourceWorkbook = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
    End If
    sourceTab = "Clients"
    
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
    recSet.Source = "SELECT * FROM [" & sourceTab & "$]"
    recSet.Open
    
    'Copy to wshBD_Clients workbook
    wshClients.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet - 2024-07-20 @ 18:31
    Dim rng As Range: Set rng = wshClients.Range("A1").CurrentRegion
    Call Apply_Worksheet_Format(wshClients, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""

    'Cleaning memory - 2024-07-01 @ 09:34
    Set connStr = Nothing
    Set recSet = Nothing
    
'    MsgBox "Temps requis = " & Timer - timerStart
        
    Call Log_Record("modMain:Client_List_Import_All", startTime)

End Sub

Sub Apply_Worksheet_Format(ws As Worksheet, rng As Range, headerRow As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Apply_Worksheet_Format", 0)
    
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
    '        usedRange.FormatConditions.add Type:=xlExpression, _
    '            Formula1:="=ET($A2<>"""";MOD(LIGNE();2)=1)"
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
    
    Call Log_Record("modMain:Apply_Worksheet_Format", startTime)

End Sub

Function Fn_Does_Client_Code_Exist() As Boolean

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Fn_Does_Client_Code_Exist", 0)
    
    Fn_Does_Client_Code_Exist = False
    
    Dim ws As Worksheet: Set ws = wshClients
    Dim iCodeClient As String
    iCodeClient = frmForm.txtCodeClient.Value
    If iCodeClient = "" Then
        Exit Function
    End If
    
    'Validating Duplicate Entries
    If Not ws.Range("B:B").Find(what:=iCodeClient, lookat:=xlWhole) Is Nothing Then
        Fn_Does_Client_Code_Exist = True
    End If

    Call Log_Record("modMain:Fn_Does_Client_Code_Exist", startTime)

End Function

Function Fix_Txt_Fin_Annee(fyem As String) As String

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Fix_Txt_Fin_Annee", 0)
    
    'Add the last day of the month to Fiscal Year end month
    Dim fiscalYearEndString As String
    Select Case fyem
        Case "Janvier"
            fiscalYearEndString = "31/01"
        Case "Février"
            fiscalYearEndString = "28/02"
        Case "Mars"
            fiscalYearEndString = "31/03"
        Case "Avril"
            fiscalYearEndString = "30/04"
        Case "Mai"
            fiscalYearEndString = "31/05"
        Case "Juin"
            fiscalYearEndString = "30/06"
        Case "Juillet"
            fiscalYearEndString = "31/07"
        Case "Août"
            fiscalYearEndString = "31/08"
        Case "Septembre"
            fiscalYearEndString = "30/09"
        Case "Octobre"
            fiscalYearEndString = "31/10"
        Case "Novembre"
            fiscalYearEndString = "30/11"
        Case "Décembre"
            fiscalYearEndString = "31/12"
        Case Else
            fiscalYearEndString = frmForm.cmbFinAnnee.Value
        End Select
        
    Fix_Txt_Fin_Annee = fiscalYearEndString

    Call Log_Record("modMain:Fix_Txt_Fin_Annee", startTime)

End Function
