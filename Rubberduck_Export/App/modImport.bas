Attribute VB_Name = "modImport"
Option Explicit

Sub ImporterMASTERGenerique(sourceWb As String, ws As Worksheet, onglet As String, table As String) '2025-05-07 @ 18:00

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterMASTERGenerique:" & onglet, vbNullString, 0)
    
'    On Error GoTo ERREUR_IMPORT
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '1. Vider la table locale
    Call ViderTableau(onglet, table)
    
    '2. Construire le chemin complet
    Dim fullPathSourceWb As String
    fullPathSourceWb = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & sourceWb

    If Dir(fullPathSourceWb) = "" Then
        MsgBox "Fichier source introuvable : " & fullPathSourceWb, _
        vbCritical, _
        "Le fichier source (MASTER) est introuvable"
        Call EnregistrerErreurs("modImport", "ImporterMASTERGenerique", "Ouverture", Err.Number)
        GoTo FIN
    End If
    
    '3. Connexion ADO
    Dim conn As ADODB.Connection: Set conn = New ADODB.Connection
    Dim t0 As Double: t0 = Timer
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullPathSourceWb & ";" & _
              "Extended Properties='Excel 12.0 Xml;HDR=YES';"

    '4. Lecture ADO
    Dim recSet As ADODB.Recordset: Set recSet = New ADODB.Recordset
    Dim sourceTab As String: sourceTab = "[" & onglet & "$]"
    recSet.Open "SELECT * FROM " & sourceTab, conn, adOpenStatic, adLockReadOnly

    Debug.Print vbNullString
    Debug.Print Now() & " Importation générique pour " & sourceTab & " - recSet.State       = " & recSet.state
    Debug.Print Now() & " Importation générique pour " & sourceTab & " - recSet.EOF         = " & recSet.EOF
    Debug.Print Now() & " Importation générique pour " & sourceTab & " - recSet.RecordCount = " & recSet.RecordCount
    Debug.Print Now() & " Importation générique pour " & sourceTab & " - Temps requis       = " & Format$(Timer - startTime, "0.0000 secondes")
    
    '5. Injection directe dans la table structurée
    Dim tbl As ListObject: Set tbl = ws.ListObjects(table)
    If tbl Is Nothing Then
        MsgBox "Table Excel introuvable : " & table, _
            vbCritical, _
            "Impossible de trouver de '" & table & "'"
        Call EnregistrerErreurs("modImport", "ImporterMASTERGenerique", "Impossible de trouver de '" & _
                                table & "'", Err.Number)
        GoTo FIN
    End If

    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.ClearContents
        tbl.DataBodyRange.Cells(1, 1).CopyFromRecordset recSet
    Else
        ' Table vide : injecter juste sous l’en-tête
        tbl.HeaderRowRange.offset(1, 0).CopyFromRecordset recSet
    End If
    
    '6. Audit et format
    If ws.AutoFilterMode Then ws.ShowAllData
    Call AppliquerFormatColonnesParTable(ws, ws.Range("A1").CurrentRegion, tbl.HeaderRowRange.row)
    
FIN:
    '7. Nettoyage
    If Not recSet Is Nothing Then If recSet.state <> 0 Then recSet.Close
    If Not conn Is Nothing Then If conn.state <> 0 Then conn.Close
    Set conn = Nothing
    Set recSet = Nothing
    Set tbl = Nothing

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterMASTERGenerique:" & onglet, vbNullString, startTime)
    Exit Sub
    
ERREUR_IMPORT:
    Call EnregistrerErreurs("modImport", "ImporterMASTERGenerique", sourceTab, Err.Number, "ERREUR")
    MsgBox "Erreur [" & Err.Number & "] : " & Err.description, _
        vbCritical, _
        "Importation de " & sourceTab
    Resume FIN

End Sub

Sub ImporterClients() 'Using ADODB - 2024-02-25 @ 10:23
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterClients", vbNullString, 0)
    
    Application.StatusBar = "Importation des Clients à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    '1. Vider la table locale
    Dim strFeuille As String: strFeuille = "BD_Clients"
    Dim strTable As String: strTable = "l_tbl_BD_Clients"
    Call ViderTableau(strFeuille, strTable)
    
    '2. Importer les enregistrements de GCF_MASTER.xlsx
    Dim ws As Worksheet: Set ws = wsdBD_Clients
    
    'Import Clients List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                     wsdADMIN.Range("CLIENTS_FILE").Value
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
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterClients", vbNullString, startTime)

End Sub

Sub ImporterDebRecurrent() '2025-05-07 @ 14:14
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterDebRecurrent", vbNullString, 0)
    
    Application.StatusBar = "Importation des Déboursés récurrents à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdDEB_Recurrent
    Dim onglet As String: onglet = "DEB_Recurrent"
    Dim table As String: table = "l_tbl_DEB_Recurrent"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterDebRecurrent", vbNullString, startTime)

End Sub

Sub ImporterDebTrans() '2025-05-07 @ 14:25
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterDebTrans", vbNullString, 0)
    
    Application.StatusBar = "Importation des déboursés à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdDEB_Trans
    Dim onglet As String: onglet = "DEB_Trans"
    Dim table As String: table = "l_tbl_DEB_Trans"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterDebTrans", vbNullString, startTime)

End Sub

Sub ImporterEncDetails() '2025-05-07 @ 14:45
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEncDetails", vbNullString, 0)
    
    Application.StatusBar = "Importation du détail des encaissements à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdENC_Details
    Dim onglet As String: onglet = "ENC_Details"
    Dim table As String: table = "l_tbl_ENC_Details"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEncDetails", vbNullString, startTime)

End Sub

Sub ImporterEncEntete() '2025-05-07 @ 14:50
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEncEntete", vbNullString, 0)
    
    Application.StatusBar = "Importation des encaissements à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdENC_Entete
    Dim onglet As String: onglet = "ENC_Entete"
    Dim table As String: table = "l_tbl_ENC_Entete"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEncEntete", vbNullString, startTime)

End Sub

Sub ImporterCCRegularisations() '2025-05-07 @ 13:58
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterCCRegularisations", vbNullString, 0)
    
    Application.StatusBar = "Importation des Régularisations à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdCC_Regularisations
    Dim onglet As String: onglet = "CC_Regularisations"
    Dim table As String: table = "l_tbl_CC_Regularisations"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterCCRegularisations", vbNullString, startTime)

End Sub

Sub ImporterFacComptesClients() '2025-05-07 @ 14:52
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacComptesClients", vbNullString, 0)
    
    Application.StatusBar = "Importation des comptes-clients à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdFAC_Comptes_Clients
    Dim onglet As String: onglet = "FAC_Comptes_Clients"
    Dim table As String: table = "l_tbl_FAC_Comptes_Clients"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacComptesClients", vbNullString, startTime)

End Sub

Sub ImporterFacDetails() '2025-05-07 @ 14:59
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacDetails", vbNullString, 0)
    
    Application.StatusBar = "Importation des détails de Facture à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdFAC_Details
    Dim onglet As String: onglet = "FAC_Details"
    Dim table As String: table = "l_tbl_FAC_Details"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacDetails", vbNullString, startTime)

End Sub

Sub ImporterFacEntete() '2025-05-07 @ 15:02
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacEntete", vbNullString, 0)
    
    Application.StatusBar = "Importation des entêtes de Facture à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    Dim onglet As String: onglet = "FAC_Entete"
    Dim table As String: table = "l_tbl_FAC_Entete"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacEntete", vbNullString, startTime)

End Sub

Sub ImporterFacSommaireTaux() '2025-05-07 @ 16:08
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacSommaireTaux", vbNullString, 0)
    
    Application.StatusBar = "Importation des sommaire de taux (facture) à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdFAC_Sommaire_Taux
    Dim onglet As String: onglet = "FAC_Sommaire_Taux"
    Dim table As String: table = "l_tbl_FAC_Sommaire_Taux"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacSommaireTaux", vbNullString, startTime)

End Sub

Sub ImporterFacProjetsDetails() '2025-05-07 @ 15:57
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacProjetsDetails", vbNullString, 0)
    
    Application.StatusBar = "Importation du détail des projets de facture à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
        
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Details
    Dim onglet As String: onglet = "FAC_Projets_Details"
    Dim table As String: table = "l_tbl_FAC_Projets_Details"

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
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set lo = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacProjetsDetails", vbNullString, startTime)

End Sub

Sub ImporterFacProjetsEntete() '2025-05-07 @ 16:05
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacProjetsEntete", vbNullString, 0)
    
    Application.StatusBar = "Importation des entêtes de projets de facture à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Entete
    Dim onglet As String: onglet = "FAC_Projets_Entete"
    Dim table As String: table = "l_tbl_FAC_Projets_Entete"

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
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set lo = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFacProjetsEntete", vbNullString, startTime)

End Sub

Sub ImporterFournisseurs() 'Using ADODB - 2024-07-03 @ 15:43
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFournisseurs", vbNullString, 0)
    
    Application.StatusBar = "Importation des Fournisseurs à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Clear all cells, but the headers, in the destination worksheet
    wsdBD_Fournisseurs.Range("A1").CurrentRegion.offset(1, 0).ClearContents

    'Import Suppliers List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                     wsdADMIN.Range("CLIENTS_FILE").Value
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
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set conn = Nothing
    Set recSet = Nothing
    Set rng = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterFournisseurs", vbNullString, startTime)

End Sub

Sub ImporterEJRecurrente() '2025-05-07 @ 14:35

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEJRecurrente", vbNullString, 0)
    
    Application.StatusBar = "Importation des écritures récurrentes à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdGL_EJ_Recurrente
    Dim onglet As String: onglet = "GL_EJ_Recurrente"
    Dim table As String: table = "l_tbl_GL_EJ_Auto"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Call ConstruireSommaireEJRecurrente
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterEJRecurrente", vbNullString, startTime)

End Sub

Sub ImporterGLTransactions() '2025-05-07 @ 16:10
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterGLTransactions", vbNullString, 0)
    
    Application.StatusBar = "Importation des transactions du G/L à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdGL_Trans
    Dim onglet As String: onglet = "GL_Trans"
    Dim table As String: table = "l_tbl_GL_Trans"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterGLTransactions", vbNullString, startTime)

End Sub

Sub ImporterTEC()
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterTEC", vbNullString, 0)
    
    Application.StatusBar = "Importation des TEC à partir de GCF_MASTER.xlsx"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String: sourceWb = wsdADMIN.Range("MASTER_FILE").Value
    Dim ws As Worksheet: Set ws = wsdTEC_Local
    Dim onglet As String: onglet = "TEC_Local"
    Dim table As String: table = "l_tbl_TEC_Local"

    Call ImporterMASTERGenerique(sourceWb, ws, onglet, table)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modImport:ImporterTEC", vbNullString, startTime)

End Sub

Sub ViderTableau(nomFeuille As String, nomTableau As String) '2025-05-07 @ 10:13
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomFeuille)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(nomTableau)

    On Error Resume Next '2025-08-11 @ 08:49
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If
    On Error GoTo 0
    
    'Libérer la mémoire
    Set tbl = Nothing
    Set ws = Nothing

End Sub


