Attribute VB_Name = "modImport"
Option Explicit

Sub ImporterPlanComptable() '2024-02-17 @ 07:21

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterPlanComptable", "", 0)

    'Clear all cells, but the headers, in the target worksheet
    wsdADMIN.Range("T10").CurrentRegion.offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wsdADMIN.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx"
    sourceWorksheet = "PlanComptable$"

    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection

    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open

    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = "SELECT * FROM [" & sourceWorksheet & "]"
        .Open
    End With

    'Copy to wsdAdmin workbook
    wsdADMIN.Range("T11").CopyFromRecordset recSet

    'Close resource
    recSet.Close
    connStr.Close

    Call RedefinirDnrPlanComptable

    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing

    Call Log_Record("modImport:ImporterPlanComptable", "", startTime)

End Sub

Sub ImportGeneriqueDuMaster(sourceWb As String, ws As Worksheet, onglet As String, table As String) '2025-05-07 @ 18:00

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImportGeneriqueDuMaster:" & onglet, "", 0)
    
    Application.ScreenUpdating = False
    
    '1. Vider la table locale
    Call ViderTableau(onglet, table)
    
    '2. Importer les enregistrements de la source
    Dim fullPathSourceWb As String, sourceTab As String
    fullPathSourceWb = wsdADMIN.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                       sourceWb
    sourceTab = onglet & "$"
                     
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & fullPathSourceWb & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Utilisation de la table de la feuille
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
    
'    tbl.tableStyle = "TableStyleMedium2"
    If tbl.ShowTableStyleRowStripes = False Then tbl.ShowTableStyleRowStripes = True
    
    Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(ws, rng, tbl.HeaderRowRange.row)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set rng = Nothing
    Set targetCell = Nothing
    Set tbl = Nothing
    
    Call Log_Record("modImport:ImportGeneriqueDuMaster:" & onglet, "", startTime)

End Sub

Sub ImporterClients() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterClients", "", 0)
    
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
    sourceWorkbook = wsdADMIN.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx" '2024-02-14 @ 07:04
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
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copier le recSet vers ws
    If recSet.EOF = False Then
        ws.Range("A2").CopyFromRecordset recSet
    End If
    
    'Close resource
    recSet.Close
    connStr.Close
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterClients", "", startTime)

End Sub

Sub ImporterDebRecurrent() '2025-05-07 @ 14:14
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterDebRecurrent", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdDEB_Recurrent
    Dim onglet As String, table As String
    onglet = "DEB_Récurrent"
    table = "l_tbl_DEB_Recurrent"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterDebRecurrent", "", startTime)

End Sub

Sub ImporterDebTrans() '2025-05-07 @ 14:25
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterDebTrans", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdDEB_Trans
    Dim onglet As String, table As String
    onglet = "DEB_Trans"
    table = "l_tbl_DEB_Trans"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterDebTrans", "", startTime)

End Sub

Sub ImporterEncDetails() '2025-05-07 @ 14:45
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterEncDetails", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdENC_Details
    Dim onglet As String, table As String
    onglet = "ENC_Détails"
    table = "l_tbl_ENC_Détails"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterEncDetails", "", startTime)

End Sub

Sub ImporterEncEntete() '2025-05-07 @ 14:50
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterEncEntete", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdENC_Entete
    Dim onglet As String, table As String
    onglet = "ENC_Entête"
    table = "l_tbl_ENC_Entête"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterEncEntete", "", startTime)

End Sub

Sub ImporterCCRegularisations() '2025-05-07 @ 13:58
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterCCRegularisations", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdCC_Regularisations
    Dim onglet As String, table As String
    onglet = "CC_Régularisations"
    table = "l_tbl_CC_Régularisations"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterCCRegularisations", "", startTime)

End Sub

Sub ImporterFacComptesClients() '2025-05-07 @ 14:52
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacComptesClients", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdFAC_Comptes_Clients
    Dim onglet As String, table As String
    onglet = "FAC_Comptes_Clients"
    table = "l_tbl_FAC_Comptes_Clients"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterFacComptesClients", "", startTime)

End Sub

Sub ImporterFacDetails() '2025-05-07 @ 14:59
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacDetails", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdFAC_Details
    Dim onglet As String, table As String
    onglet = "FAC_Détails"
    table = "l_tbl_FAC_Détails"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterFacDetails", "", startTime)

End Sub

Sub ImporterFacEntete() '2025-05-07 @ 15:02
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacEntete", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdFAC_Entete
    Dim onglet As String, table As String
    onglet = "FAC_Entête"
    table = "l_tbl_FAC_Entête"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterFacEntete", "", startTime)

End Sub

Sub ImporterFacSommaireTaux() '2025-05-07 @ 16:08
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacSommaireTaux", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdFAC_Sommaire_Taux
    Dim onglet As String, table As String
    onglet = "FAC_Sommaire_Taux"
    table = "l_tbl_FAC_Sommaire_Taux"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterFacSommaireTaux", "", startTime)

End Sub

Sub ImporterFacProjetsDetails() '2025-05-07 @ 15:57
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacProjetsDetails", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdFAC_Projets_Details
    Dim onglet As String, table As String
    onglet = "FAC_Projets_Détails"
    table = "l_tbl_FAC_Projets_Détails"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterFacProjetsDetails", "", startTime)

End Sub

Sub ImporterFacProjetsEntete() '2025-05-07 @ 16:05
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFacProjetsEntete", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdFAC_Projets_Entete
    Dim onglet As String, table As String
    onglet = "FAC_Projets_Entête"
    table = "l_tbl_FAC_Projets_Entête"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterFacProjetsEntete", "", startTime)

End Sub

Sub ImporterFournisseurs() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterFournisseurs", "", 0)
    
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wsdBD_Fournisseurs.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx" '2024-02-14 @ 07:04
    sourceTab = "Fournisseurs$"
    
    'ADODB connection
    Dim connStr As ADODB.Connection: Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    With recSet
        .ActiveConnection = connStr
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = "SELECT * FROM [" & sourceTab & "]"
        .Open
    End With
    
    'Copy to wsdBD_Fournisseurs workbook
    wsdBD_Fournisseurs.Range("A2").CopyFromRecordset recSet
    
    'Setup the format of the worksheet using a Sub - 2024-07-20 @ 18:38
    Dim rng As Range: Set rng = wsdBD_Fournisseurs.Range("A1").CurrentRegion
    Call AppliquerFormatColonnesParTable(wsdBD_Fournisseurs, rng, 1)
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set connStr = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call Log_Record("modImport:ImporterFournisseurs", "", startTime)

End Sub

Sub ImporterEJRecurrente() '2025-05-07 @ 14:35

    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterEJRecurrente", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdGL_EJ_Recurrente
    Dim onglet As String, table As String
    onglet = "GL_EJ_Récurrente"
    table = "l_tbl_GL_EJ_Auto"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    Call GL_EJ_Recurrente_Build_Summary '2024-03-14 @ 07:38
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterEJRecurrente", "", startTime)

End Sub

Sub ImporterGLTransactions() '2025-05-07 @ 16:10
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterGLTransactions", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdGL_Trans
    Dim onglet As String, table As String
    onglet = "GL_Trans"
    table = "l_tbl_GL_Trans"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterGLTransactions", "", startTime)

End Sub

Sub ImporterTEC() '2024-02-14 @ 06:19
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterTEC", "", 0)
    
    Application.StatusBar = "Importation des TEC à partir de GCF_MASTER.xlsx" '2025-06-13 @ 08:47
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdTEC_Local
    Dim onglet As String, table As String
    onglet = "TEC_Local"
    table = "l_tbl_TEC_Local"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Application.StatusBar = False '2025-06-13 @ 08:47
    
    Call Log_Record("modImport:ImporterTEC", "", startTime)

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

