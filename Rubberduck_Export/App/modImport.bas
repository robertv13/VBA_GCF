Attribute VB_Name = "modImport"
Option Explicit

Sub ImporterMASTERGenerique(sourceWb As String, ws As Worksheet, onglet As String, table As String) '2025-05-07 @ 18:00

'    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterMASTERGenerique:" & onglet, vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    '1. Vider la table locale
    Call ViderTableau(onglet, table)
    
    '2. Importer les enregistrements de la source via ADO
    Dim fullPathSourceWb As String, sourceTab As String
    fullPathSourceWb = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                       sourceWb
    sourceTab = onglet & "$"
                     
    'ADODB connection
    Dim conn As Object: Set conn = New ADODB.Connection
    
    'Connection String specific to EXCEL
    conn.Open = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source = " & fullPathSourceWb & ";" & _
                "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & sourceTab & "]"
    With recSet
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = strSQL
        .Open
    End With
    
    'Déclaration du tableau structuré de la feuille
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(table)
    
    'Copy to local worksheet
    Dim targetCell As Range
    If Not recSet.EOF Then
        If Not tbl.DataBodyRange Is Nothing Then
            'Si la table a déjà des lignes, on remplace à partir de la première
            Set targetCell = tbl.DataBodyRange.Cells(1, 1)
        Else
            'Si la table est vide, on commence une ligne après l’en-tête
            Set targetCell = tbl.HeaderRowRange.offset(1, 0).Cells(1, 1)
        End If
        targetCell.CopyFromRecordset recSet
    End If
    
    Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(ws, rng, tbl.HeaderRowRange.row)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set rng = Nothing
    Set targetCell = Nothing
    Set tbl = Nothing
    
'    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterMASTERGenerique:" & onglet, vbNullString, startTime)

End Sub

Sub ImporterClients() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterClients", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    '1. Vider la table locale
    Dim strFeuille As String
    strFeuille = "BD_Clients"
    Dim strTable As String
    strTable = "l_tbl_BD_Clients"
    Call ViderTableau(strFeuille, strTable)
    
    '2. Importer les enregistrements de GCF_MASTER.xlsx
    Dim ws As Worksheet
    Set ws = wsdBD_Clients
    
    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                     wsdADMIN.Range("CLIENTS_FILE").Value '2024-02-14 @ 07:04
    sourceTab = "Clients$"
    
    'ADODB connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source = " & sourceWorkbook & ";" & _
                "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")
    
    With recSet
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copier le recSet vers ws
    If recSet.EOF = False Then
        ws.Range("A2").CopyFromRecordset recSet
    End If
    
    'Close resource
    recSet.Close
    conn.Close
    
    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterClients", vbNullString, startTime)

End Sub

Sub ImporterDebRecurrent() '2025-05-07 @ 14:14
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterDebRecurrent", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdDEB_Recurrent
    Dim onglet As String, table As String
    onglet = "DEB_Recurrent"
    table = "l_tbl_DEB_Recurrent"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterDebRecurrent", vbNullString, startTime)

End Sub

Sub ImporterDebTrans() '2025-05-07 @ 14:25
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterDebTrans", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdDEB_Trans
    Dim onglet As String, table As String
    onglet = "DEB_Trans"
    table = "l_tbl_DEB_Trans"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterDebTrans", vbNullString, startTime)

End Sub

Sub ImporterEncDetails() '2025-05-07 @ 14:45
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEncDetails", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdENC_Details
    Dim onglet As String, table As String
    onglet = "ENC_Details"
    table = "l_tbl_ENC_Details"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEncDetails", vbNullString, startTime)

End Sub

Sub ImporterEncEntete() '2025-05-07 @ 14:50
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEncEntete", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdENC_Entete
    Dim onglet As String, table As String
    onglet = "ENC_Entete"
    table = "l_tbl_ENC_Entete"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEncEntete", vbNullString, startTime)

End Sub

Sub ImporterCCRegularisations() '2025-05-07 @ 13:58
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterCCRegularisations", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdCC_Regularisations
    Dim onglet As String, table As String
    onglet = "CC_Regularisations"
    table = "l_tbl_CC_Regularisations"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterCCRegularisations", vbNullString, startTime)

End Sub

Sub ImporterFacComptesClients() '2025-05-07 @ 14:52
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacComptesClients", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdFAC_Comptes_Clients
    Dim onglet As String, table As String
    onglet = "FAC_Comptes_Clients"
    table = "l_tbl_FAC_Comptes_Clients"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacComptesClients", vbNullString, startTime)

End Sub

Sub ImporterFacDetails() '2025-05-07 @ 14:59
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacDetails", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdFAC_Details
    Dim onglet As String, table As String
    onglet = "FAC_Details"
    table = "l_tbl_FAC_Details"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacDetails", vbNullString, startTime)

End Sub

Sub ImporterFacEntete() '2025-05-07 @ 15:02
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacEntete", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdFAC_Entete
    Dim onglet As String, table As String
    onglet = "FAC_Entete"
    table = "l_tbl_FAC_Entete"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacEntete", vbNullString, startTime)

End Sub

Sub ImporterFacSommaireTaux() '2025-05-07 @ 16:08
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacSommaireTaux", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdFAC_Sommaire_Taux
    Dim onglet As String, table As String
    onglet = "FAC_Sommaire_Taux"
    table = "l_tbl_FAC_Sommaire_Taux"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacSommaireTaux", vbNullString, startTime)

End Sub

Sub ImporterFacProjetsDetails() '2025-05-07 @ 15:57
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacProjetsDetails", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdFAC_Projets_Details
    Dim onglet As String, table As String
    onglet = "FAC_Projets_Details"
    table = "l_tbl_FAC_Projets_Details"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Enlever la ligne fantôme associé à ADO avec fichier source vide... 2025-07-09 @ 07:27
    Dim ligne2Vide As Boolean
    Dim donneesEnLignes3EtPlus As Boolean
    
    ligne2Vide = (Application.WorksheetFunction.CountA(ws.Rows(2)) = 0)
    donneesEnLignes3EtPlus = (Application.WorksheetFunction.CountA(ws.Range("3:" & ws.Rows.count)) > 0)
    
    If ligne2Vide And donneesEnLignes3EtPlus Then
        Dim lo As ListObject
        Set lo = ws.ListObjects(1)
        If Application.WorksheetFunction.CountA(ws.Rows(2)) = 0 Then
            'Suppression via la table directement
            lo.ListRows(1).Delete
        End If
    End If
    
    'Libérer la mémoire
    Set lo = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacProjetsDetails", vbNullString, startTime)

End Sub

Sub ImporterFacProjetsEntete() '2025-05-07 @ 16:05
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacProjetsEntete", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdFAC_Projets_Entete
    Dim onglet As String, table As String
    onglet = "FAC_Projets_Entete"
    table = "l_tbl_FAC_Projets_Entete"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Enlever la ligne fantôme associé à ADO avec fichier source vide... 2025-07-09 @ 07:46
    Dim ligne2Vide As Boolean
    Dim donneesEnLignes3EtPlus As Boolean
    
    ligne2Vide = (Application.WorksheetFunction.CountA(ws.Rows(2)) = 0)
    donneesEnLignes3EtPlus = (Application.WorksheetFunction.CountA(ws.Range("3:" & ws.Rows.count)) > 0)
    
    If ligne2Vide And donneesEnLignes3EtPlus Then
        Dim lo As ListObject
        Set lo = ws.ListObjects(1)
        If Application.WorksheetFunction.CountA(ws.Rows(2)) = 0 Then
            'Suppression via la table directement
            lo.ListRows(1).Delete
        End If
    End If
    
    'Libérer la mémoire
    Set lo = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacProjetsEntete", vbNullString, startTime)

End Sub

Sub ImporterFournisseurs() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFournisseurs", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wsdBD_Fournisseurs.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                     wsdADMIN.Range("CLIENTS_FILE").Value '2024-02-14 @ 07:04
    sourceTab = "Fournisseurs$"
    
    'ADODB connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source = " & sourceWorkbook & ";" & _
                "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")
    
    With recSet
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wsdBD_Fournisseurs workbook
    wsdBD_Fournisseurs.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    Dim rng As Range: Set rng = wsdBD_Fournisseurs.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdBD_Fournisseurs, rng, 1)
    
    'Close resource
    recSet.Close
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFournisseurs", vbNullString, startTime)

End Sub

Sub ImporterEJRecurrente() '2025-05-07 @ 14:35

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEJRecurrente", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdGL_EJ_Recurrente
    Dim onglet As String, table As String
    onglet = "GL_EJ_Recurrente"
    table = "l_tbl_GL_EJ_Auto"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Call ConstruireSommaireEJRecurrente '2024-03-14 @ 07:38
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEJRecurrente", vbNullString, startTime)

End Sub

Sub ImporterGLTransactions() '2025-05-07 @ 16:10
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterGLTransactions", vbNullString, 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdGL_Trans
    Dim onglet As String, table As String
    onglet = "GL_Trans"
    table = "l_tbl_GL_Trans"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterGLTransactions", vbNullString, startTime)

End Sub

Sub ImporterTEC() '2024-02-14 @ 06:19
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterTEC", vbNullString, 0)
    
    Application.StatusBar = "Importation des TEC à partir de GCF_MASTER.xlsx" '2025-06-13 @ 08:47
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet
    Set ws = wsdTEC_Local
    Dim onglet As String, table As String
    onglet = "TEC_Local"
    table = "l_tbl_TEC_Local"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Application.StatusBar = False '2025-06-13 @ 08:47
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterTEC", vbNullString, startTime)

End Sub

Sub ViderTableau(nomFeuille As String, nomTableau As String) '2025-05-07 @ 10:13
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomFeuille)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(nomTableau)

    If tbl.ListRows.count > 0 Then
        tbl.DataBodyRange.Delete
    End If

End Sub


