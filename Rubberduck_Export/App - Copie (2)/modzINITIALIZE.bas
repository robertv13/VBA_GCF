Attribute VB_Name = "modzINITIALIZE"
Option Explicit

Sub DeleteAllRowsButKeepHeaders() '2024-07-30 @ 12:21
    
    '*************************************************************************************
    'This procedure delete all rows from all the tables
    'Use only to install a clean version of the application
    'It's also reset any counters (invoice number, transaction number, etc.)
    '*************************************************************************************

    'Step 1 - Erase all rows from GCF_BD_MASTER
    
    'Define workbook path
    Dim sourcePath As String
    sourcePath = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                "GCF_BD_MASTER.xlsx" '2024-07-29 @ 18:17

    'Ouvrir le workbook
    Dim wb As Workbook: Set wb = Workbooks.Open(sourcePath)

    'Boucle à travers chaque feuille
    Dim ws As Worksheet
    Dim lastUsedRow As Long
    For Each ws In wb.Worksheets
        ' Trouver la dernière ligne utilisée
        If InStr(ws.Name, "Admin") <> 1 Then
            lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
            'Supprimer toutes les lignes sauf la première (en-tête)
            If lastUsedRow > 1 Then
                ws.Range("A2").CurrentRegion.offset(1, 0).ClearContents
            End If
        End If
    Next ws

    'Sauvegarder et fermer le workbook
    wb.Close SaveChanges:=True

    'Step 2 - Enlève les rangées d'une feuille locale
    Set ws = wshTEC_TDB_Data
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    'Supprimer toutes les lignes sauf la première (en-tête)
    If lastUsedRow > 1 Then
        ws.Range("A2").CurrentRegion.offset(1, 0).ClearContents
        
        'Requires a minimum of one line (values)
        ws.Range("A2").value = 0
        ws.Range("B2").value = ""
        ws.Range("C2").value = Format$(Date, wsdADMIN.Range("B1").value)
        ws.Range("E2").value = 0
        ws.Range("F2").value = "VRAI"
        ws.Range("G2").value = "FAUX"
        ws.Range("H2").value = "FAUX"
        'Formulas for other columns
        ws.Range("I2").formula = "=IF([@EstDetruite],E2,0)"
        ws.Range("J2").formula = "=[@[H_Saisies]]-[@[H_Détruites]]"
        ws.Range("K2").formula = "=IF([@EstFacturable]=""VRAI"",(E2-I2),0)"
        ws.Range("L2").formula = "=(E2-I2)-K2"
        ws.Range("M2").formula = "=IF(G2=""VRAI"",J2,0)"
        ws.Range("N2").formula = "=MAX([@[H_Facturables]]-[@[H_Facturées]],0)"
        ws.Range("O2").formula = "=ROUND(INT(NOW()-[@Date]),0)"
        ws.Range("P2").formula = "=IF([@Âge]<=30,N2,0)"
        ws.Range("Q2").formula = "=IF(AND([@Âge]>30,[@Âge]<=60),N2,0)"
        ws.Range("R2").formula = "=IF(AND([@Âge]>60,[@Âge]<=90),N2,0)"
        ws.Range("S2").formula = "=IF([@Âge]>90,N2,0)"
    End If
    
    'Step # 2 - Import all worksheets from Sortie.xlsx
    Call modImport.ImporterClients
    Call modImport.ImporterFournisseurs
    
    Call modImport.ImporterDebRecurrent
    Call modImport.ImporterDebTrans
    
    Call modImport.ImporterEncDetails
    Call modImport.ImporterEncEntete
    
    Call modImport.ImporterFacComptesClients
    Call modImport.ImporterFacDetails
    Call modImport.ImporterFacEntete
    Call modImport.ImporterFacProjetsDetails
    Call modImport.ImporterFacProjetsEntete
    
    Call modImport.ImporterEJRecurrente
    Call modImport.ImporterGLTransactions
    
    Call modImport.ImporterTEC
    
    'Step # 3 - Process the current workbook
    Set ws = wshDEB_Saisie
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("B1").value = 0
    ws.Range("B2").value = 0
    ws.Range("B3").value = 0
    Application.EnableEvents = True
    
    Set ws = wshFAC_Brouillon
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("B21").value = 1
    ws.Range("B33:B49").ClearContents
    ws.Range("B51").ClearContents
    ws.Range("B52").ClearContents
    ws.Range("B53").ClearContents
    ws.Range("B54").ClearContents
    Application.EnableEvents = True
    
    Set ws = wshGL_BV
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("B3").value = "31/07/2024"
    Application.EnableEvents = True
    
    Set ws = wshGL_EJ
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("B1").value = 1
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set wb = Nothing
    Set ws = Nothing
    
    MsgBox "Toutes les données ont été supprimées avec succès," & vbNewLine & vbNewLine & _
           "en gardant les en-têtes !"
    
End Sub


