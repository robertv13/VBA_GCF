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
    sourcePath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                "GCF_BD_MASTER.xlsx" '2024-07-29 @ 18:17

    'Ouvrir le workbook
    Dim wb As Workbook: Set wb = Workbooks.Open(sourcePath)

    'Boucle � travers chaque feuille
    Dim ws As Worksheet
    Dim lastUsedRow As Long
    For Each ws In wb.Worksheets
        ' Trouver la derni�re ligne utilis�e
        If InStr(ws.Name, "Admin") <> 1 Then
            lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
            'Supprimer toutes les lignes sauf la premi�re (en-t�te)
            If lastUsedRow > 1 Then
                ws.Range("A2").CurrentRegion.offset(1, 0).ClearContents
            End If
        End If
    Next ws

    'Sauvegarder et fermer le workbook
    wb.Close SaveChanges:=True

    'Step 2 - Enl�ve les rang�es d'une feuille locale
    Set ws = wshTEC_TDB_Data
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    'Supprimer toutes les lignes sauf la premi�re (en-t�te)
    If lastUsedRow > 1 Then
        ws.Range("A2").CurrentRegion.offset(1, 0).ClearContents
        
        'Requires a minimum of one line (values)
        ws.Range("A2").Value = 0
        ws.Range("B2").Value = ""
        ws.Range("C2").Value = Format$(Date, wshAdmin.Range("B1").Value)
        ws.Range("E2").Value = 0
        ws.Range("F2").Value = "VRAI"
        ws.Range("G2").Value = "FAUX"
        ws.Range("H2").Value = "FAUX"
        'Formulas for other columns
        ws.Range("I2").formula = "=IF([@EstDetruite],E2,0)"
        ws.Range("J2").formula = "=[@[H_Saisies]]-[@[H_D�truites]]"
        ws.Range("K2").formula = "=IF([@EstFacturable]=""VRAI"",(E2-I2),0)"
        ws.Range("L2").formula = "=(E2-I2)-K2"
        ws.Range("M2").formula = "=IF(G2=""VRAI"",J2,0)"
        ws.Range("N2").formula = "=MAX([@[H_Facturables]]-[@[H_Factur�es]],0)"
        ws.Range("O2").formula = "=ROUND(INT(NOW()-[@Date]),0)"
        ws.Range("P2").formula = "=IF([@�ge]<=30,N2,0)"
        ws.Range("Q2").formula = "=IF(AND([@�ge]>30,[@�ge]<=60),N2,0)"
        ws.Range("R2").formula = "=IF(AND([@�ge]>60,[@�ge]<=90),N2,0)"
        ws.Range("S2").formula = "=IF([@�ge]>90,N2,0)"
    End If
    
    'Step # 2 - Import all worksheets from Sortie.xlsx
    Call Client_List_Import_All
    Call Fournisseur_List_Import_All
    
    Call DEB_R�current_Import_All
    Call DEB_Trans_Import_All
    
    Call ENC_D�tails_Import_All
    Call ENC_Ent�te_Import_All
    
    Call FAC_Comptes_Clients_Import_All
    Call FAC_D�tails_Import_All
    Call FAC_Ent�te_Import_All
    Call FAC_Projets_D�tails_Import_All
    Call FAC_Projets_Ent�te_Import_All
    
    Call GL_EJ_Recurrente_Import_All
    Call GL_Trans_Import_All
    
    Call TEC_Import_All
    
    'Step # 3 - Process the current workbook
    Set ws = wshDEB_Saisie
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("B1").Value = 0
    ws.Range("B2").Value = 0
    ws.Range("B3").Value = 0
    Application.EnableEvents = True
    
    Set ws = wshFAC_Brouillon
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("B21").Value = 1
    ws.Range("B33:B49").ClearContents
    ws.Range("B51").ClearContents
    ws.Range("B52").ClearContents
    ws.Range("B53").ClearContents
    ws.Range("B54").ClearContents
    Application.EnableEvents = True
    
    Set ws = wshGL_BV
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("B3").Value = "31/07/2024"
    Application.EnableEvents = True
    
    Set ws = wshGL_EJ
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("B1").Value = 1
    Application.EnableEvents = True
    
    'Lib�rer la m�moire
    Set wb = Nothing
    Set ws = Nothing
    
    MsgBox "Toutes les donn�es ont �t� supprim�es avec succ�s," & vbNewLine & vbNewLine & _
           "en gardant les en-t�tes !"
    
End Sub


