Attribute VB_Name = "modCC_Régularisation"
Option Explicit

'Variables globales pour le module
Public regulNo As Long
Public nextJENo As Long

Sub MAJ_Regularisation() '2025-01-14 @ 12:00
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Régularisation:MAJ_Regularisation", "", 0)
    
    With wshENC_Saisie
        'As-t-on les champs obligatoires ?
        If .Range("F5").value = Empty Or _
           .Range("K5").value = Empty Or _
           .Range("F7").value = Empty Or _
           .Range("K7").value = 0 Then
            msgBox "Assurez-vous d'avoir..." & vbNewLine & vbNewLine & _
                "1. Un client valide" & vbNewLine & _
                "2. Une date de régularisation" & vbNewLine & _
                "3. Un type de transaction et" & vbNewLine & _
                "4. Le montant de la régularisation" & vbNewLine & vbNewLine & _
                "AVANT de tenter de sauvegarder la régularisation.", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Le montant de la régularisation doit être appliqué intégralement
        If .Range("K9").value <> CCur(ufEncRégularisation.txtTotalFacture) Then
            msgBox "Assurez-vous que le montant de la régularisation soit ÉGAL" & vbNewLine & _
                "a bel et bien été réparti", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Create record to CC_Régularisations
        Call REGUL_Add_DB
        Call REGUL_Add_Locally
        
        'Update FAC_Comptes_Clients
        Call REGUL_Update_Comptes_Clients_DB
        Call REGUL_Update_Comptes_Clients_Locally
                
        'Prepare G/L posting
        Dim noRegul As String, nomClient As String, descRegul As String
        Dim dateRegul As Date
        Dim montantRegul As Currency
        dateRegul = wshENC_Saisie.Range("K5").value
        nomClient = wshENC_Saisie.Range("F5").value
        descRegul = wshENC_Saisie.Range("F9").value

        Call REGUL_GL_Posting_DB(regulNo, dateRegul, nomClient, descRegul)
        Call REGUL_GL_Posting_Locally(regulNo, dateRegul, nomClient, descRegul)
        
        msgBox "La régularisation '" & regulNo & "' a été enregistré avec succès", vbOKOnly + vbInformation, "Confirmation de traitement"
        
        'Fermer le UserForm
        Unload ufEncRégularisation
    
        Call Régularisation_Add_New 'Reset the form
        
        Call AjusteLibelléEncaissement("Banque")
    
        .Range("K5").value = Format$(Date, wshAdmin.Range("B1").value)
        .Range("B1").Select

        'De retour à la saisie du client
        .Range("F5").Select
        
    End With
    
Clean_Exit:

    Call Log_Record("modCC_Régularisation:MAJ_Regularisation", "", startTime)

End Sub

Sub Régularisation_Add_New() '2024-08-21 @ 14:58

    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Régularisation:Régularisation_Add_New", "", 0)

    Call ENC_Clear_Cells
    
    Call Log_Record("modCC_Régularisation:Régularisation_Add_New", "", startTime)
    
End Sub

Sub REGUL_Add_DB() 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Régularisation:REGUL_Add_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "CC_Régularisations$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxRegulNo As Long
    strSQL = "SELECT MAX(RegulID) AS MaxRegulNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxPmtNo
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lr As Long
    If IsNull(rs.Fields("MaxRegulNo").value) Then
        lr = 0
    Else
        lr = rs.Fields("MaxRegulNo").value
    End If
    
    'Calculate the new PmtNo
    regulNo = lr + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
        rs.Fields(fREGULRegulID - 1).value = regulNo
        rs.Fields(fREGULInvNo - 1).value = ufEncRégularisation.cbbNoFacture
        rs.Fields(fREGULDate - 1).value = CDate(wshENC_Saisie.Range("K5").value)
        rs.Fields(fREGULClientID - 1).value = wshENC_Saisie.clientCode
        rs.Fields(fREGULClientNom - 1).value = wshENC_Saisie.Range("F5").value
        rs.Fields(fREGULHono - 1).value = CCur(ufEncRégularisation.txtHonoraires)
        rs.Fields(fREGULFrais - 1).value = CCur(ufEncRégularisation.txtFraisDivers)
        rs.Fields(fREGULTPS - 1).value = CCur(ufEncRégularisation.txtTPS)
        rs.Fields(fREGULTVQ - 1).value = CCur(ufEncRégularisation.txtTVQ)
        rs.Fields(fREGULDescription - 1).value = wshENC_Saisie.Range("F9").value
        rs.Fields(fREGULTimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    'Update the recordset (create the record)
    rs.Update
    
    'Close recordset and connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modCC_Régularisation:REGUL_Add_DB", "", startTime)
    
End Sub

Sub REGUL_Add_Locally() '2024-08-22 @ 10:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Régularisation:REGUL_Add_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentRegulNo As Long
    currentRegulNo = regulNo
    
    'What is the last used row in CC_Régularisations ?
    Dim ws As Worksheet: Set ws = wshCC_Régularisations
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    ws.Range("A" & rowToBeUsed).value = regulNo
    ws.Range("B" & rowToBeUsed).value = ufEncRégularisation.cbbNoFacture
    ws.Range("C" & rowToBeUsed).value = CDate(wshENC_Saisie.Range("K5").value)
    ws.Range("D" & rowToBeUsed).value = wshENC_Saisie.clientCode
    ws.Range("E" & rowToBeUsed).value = wshENC_Saisie.Range("F5").value
    ws.Range("F" & rowToBeUsed).value = CCur(ufEncRégularisation.txtHonoraires)
    ws.Range("G" & rowToBeUsed).value = CCur(ufEncRégularisation.txtFraisDivers)
    ws.Range("H" & rowToBeUsed).value = CCur(ufEncRégularisation.txtTPS)
    ws.Range("I" & rowToBeUsed).value = CCur(ufEncRégularisation.txtTVQ)
    ws.Range("J" & rowToBeUsed).value = wshENC_Saisie.Range("F9").value
    ws.Range("K" & rowToBeUsed).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    Application.ScreenUpdating = True

    Call Log_Record("modCC_Régularisation:REGUL_Add_Locally", "", startTime)

End Sub

Sub REGUL_Update_Comptes_Clients_DB() 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Régularisation:REGUL_Update_Comptes_Clients_DB", "", 0)
    
    Dim errMsg As String
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Comptes_Clients$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Open the recordset for the specified invoice
    Dim Inv_No As String
    Inv_No = ufEncRégularisation.cbbNoFacture.value
    
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & Inv_No & "'"
    rs.Open strSQL, conn, 2, 3
    If Not (rs.BOF Or rs.EOF) Then
        Dim mntRegulTotal As Double
        mntRegulTotal = CDbl(ufEncRégularisation.txtHonoraires.value) + _
                        CDbl(ufEncRégularisation.txtFraisDivers.value) + _
                        CDbl(ufEncRégularisation.txtTPS.value) + _
                        CDbl(ufEncRégularisation.txtTVQ.value)

        'Mettre à jour Régularisation totale
        rs.Fields(fFacCCTotalRegul - 1).value = rs.Fields(fFacCCTotalRegul - 1).value + mntRegulTotal
        'Mettre à jour le solde de la facture
        rs.Fields(fFacCCBalance - 1).value = rs.Fields(fFacCCBalance - 1).value + mntRegulTotal
        'Mettre à jour Status
        If rs.Fields(fFacCCBalance - 1).value = 0 Then
            rs.Fields(fFacCCStatus - 1).value = "Paid"
        Else
            rs.Fields(fFacCCStatus - 1).value = "Unpaid"
        End If
        rs.Update
    Else
        msgBox "Problème avec la facture '" & Inv_No & "'" & vbNewLine & vbNewLine & _
               "Contactez le développeur SVP", vbCritical, "Impossible de trouver la facture dans Comptes_Clients"
    End If
    'Update the recordset (create the record)
    rs.Close
    
    'Close recordset and connection
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modCC_Régularisation:REGUL_Update_Comptes_Clients_DB", "", startTime)
    
End Sub

Sub REGUL_Update_Comptes_Clients_Locally() '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Régularisation:REGUL_Update_Comptes_Clients_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'Set the range to look for
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim Inv_No As String
    Inv_No = ufEncRégularisation.cbbNoFacture.value
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=Inv_No, LookIn:=xlValues, LookAt:=xlWhole)

    Dim rowToBeUpdated As Long
    If Not foundRange Is Nothing Then
        rowToBeUpdated = foundRange.row
        ws.Cells(rowToBeUpdated, fFacCCTotalRegul).value = ws.Cells(rowToBeUpdated, fFacCCTotalRegul).value + _
                                                            CCur(ufEncRégularisation.txtHonoraires.value) + _
                                                            CCur(ufEncRégularisation.txtFraisDivers.value) + _
                                                            CCur(ufEncRégularisation.txtTPS.value) + _
                                                            CCur(ufEncRégularisation.txtTVQ.value)
        ws.Cells(rowToBeUpdated, fFacCCBalance).value = ws.Cells(rowToBeUpdated, fFacCCBalance).value + _
                                                            CCur(ufEncRégularisation.txtHonoraires.value) + _
                                                            CCur(ufEncRégularisation.txtFraisDivers.value) + _
                                                            CCur(ufEncRégularisation.txtTPS.value) + _
                                                            CCur(ufEncRégularisation.txtTVQ.value)
 
        'Est-ce que le solde de la facture est à 0,00 $ ?
        If ws.Cells(rowToBeUpdated, fFacCCBalance).value = 0 Then
            ws.Cells(rowToBeUpdated, fFacCCStatus) = "Paid"
        Else
            ws.Cells(rowToBeUpdated, fFacCCStatus) = "Unpaid"
        End If
    Else
        msgBox "La facture '" & Inv_No & "' n'existe pas dans FAC_Comptes_Clients.", vbCritical
    End If
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modCC_Régularisation:REGUL_Update_Comptes_Clients_Locally", "", startTime)

End Sub

Sub REGUL_GL_Posting_DB(no As Long, dt As Date, nom As String, desc As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Régularisation:REGUL_GL_Posting_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans$"
    
    'Initialize connection, connection string, open the connection & declare rs Object
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJNo As Long
    strSQL = "SELECT MAX(NoEntrée) AS MaxEJNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new ID
'    Dim nextJENo As Long
    nextJENo = lastJE + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Crédit - Honoraires
    If ufEncRégularisation.txtHonoraires.value <> 0 Then
        rs.AddNew
            'Cnstruction des champs
            rs.Fields(fGlTNoEntrée - 1).value = nextJENo
            rs.Fields(fGlTDate - 1).value = Format$(dt, "yyyy-mm-dd")
            rs.Fields(fGlTDescription - 1).value = nom
            rs.Fields(fGlTSource - 1).value = "RÉGULARISATION:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).value = ObtenirNoGlIndicateur("Revenus de consultation")
            rs.Fields(fGlTCompte - 1).value = "Revenus de consultation"
            rs.Fields(fGlTDébit - 1).value = -CCur(ufEncRégularisation.txtHonoraires.value)
            rs.Fields(fGlTAutreRemarque - 1).value = desc
            rs.Fields(fGlTTimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rs.Update
    End If
    
    'Crédit - Frais Divers
    If ufEncRégularisation.txtFraisDivers.value <> 0 Then
        rs.AddNew
            'Cnstruction des champs
            rs.Fields(fGlTNoEntrée - 1).value = nextJENo
            rs.Fields(fGlTDate - 1).value = Format$(dt, "yyyy-mm-dd")
            rs.Fields(fGlTDescription - 1).value = nom
            rs.Fields(fGlTSource - 1).value = "RÉGULARISATION:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).value = ObtenirNoGlIndicateur("Revenus frais de poste")
            rs.Fields(fGlTCompte - 1).value = "Revenus - Frais de poste"
            rs.Fields(fGlTDébit - 1).value = -CCur(ufEncRégularisation.txtFraisDivers.value)
            rs.Fields(fGlTAutreRemarque - 1).value = desc
            rs.Fields(fGlTTimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rs.Update
    End If
    
    'Crédit - TPS
    If ufEncRégularisation.txtTPS.value <> 0 Then
        rs.AddNew
            'Cnstruction des champs
            rs.Fields(fGlTNoEntrée - 1).value = nextJENo
            rs.Fields(fGlTDate - 1).value = Format$(dt, "yyyy-mm-dd")
            rs.Fields(fGlTDescription - 1).value = nom
            rs.Fields(fGlTSource - 1).value = "RÉGULARISATION:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).value = ObtenirNoGlIndicateur("TPS Facturée")
            rs.Fields(fGlTCompte - 1).value = "TPS percues"
            rs.Fields(fGlTDébit - 1).value = -CCur(ufEncRégularisation.txtTPS.value)
            rs.Fields(fGlTAutreRemarque - 1).value = desc
            rs.Fields(fGlTTimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rs.Update
    End If
    
    'Crédit - TVQ
    If ufEncRégularisation.txtTVQ.value <> 0 Then
        rs.AddNew
            'Cnstruction des champs
            rs.Fields(fGlTNoEntrée - 1).value = nextJENo
            rs.Fields(fGlTDate - 1).value = Format$(dt, "yyyy-mm-dd")
            rs.Fields(fGlTDescription - 1).value = nom
            rs.Fields(fGlTSource - 1).value = "RÉGULARISATION:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).value = ObtenirNoGlIndicateur("TVQ Facturée")
            rs.Fields(fGlTCompte - 1).value = "TVQ percues"
            rs.Fields(fGlTDébit - 1).value = -CCur(ufEncRégularisation.txtTVQ.value)
            rs.Fields(fGlTAutreRemarque - 1).value = desc
            rs.Fields(fGlTTimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rs.Update
    End If
    
    'Débit = Total de Honoraires, Frais Divers, TPS & TVQ
    Dim regulTotal As Currency
    regulTotal = CCur(ufEncRégularisation.txtHonoraires.value) + _
                    CCur(ufEncRégularisation.txtFraisDivers.value) + _
                    CCur(ufEncRégularisation.txtTPS.value) + _
                    CCur(ufEncRégularisation.txtTVQ.value)
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields(fGlTNoEntrée - 1).value = nextJENo
        rs.Fields(fGlTDate - 1).value = Format$(dt, "yyyy-mm-dd")
        rs.Fields(fGlTDescription - 1).value = nom
        rs.Fields(fGlTSource - 1).value = "RÉGULARISATION:" & Format$(no, "00000")
        rs.Fields(fGlTNoCompte - 1).value = ObtenirNoGlIndicateur("Comptes Clients")
        rs.Fields(fGlTCompte - 1).value = "Comptes clients"
        rs.Fields(fGlTCrédit - 1).value = -regulTotal
        rs.Fields(fGlTAutreRemarque - 1).value = desc
        rs.Fields(fGlTTimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    rs.Update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modCC_Régularisation:REGUL_GL_Posting_DB", "", startTime)

End Sub

Sub REGUL_GL_Posting_Locally(no As Long, dt As Date, nom As String, desc As String)  'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Régularisation:REGUL_GL_Posting_Locally", "", 0)

    Application.ScreenUpdating = False

    'What is the last used row in GL_Trans ?
    Dim ws As Worksheet: Set ws = wshGL_Trans
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    rowToBeUsed = lastUsedRow + 1

    With ws
        'Credit side - Honoraires
        If ufEncRégularisation.txtHonoraires.value <> 0 Then
            .Range("A" & rowToBeUsed).value = nextJENo
            .Range("B" & rowToBeUsed).value = CDate(dt)
            .Range("C" & rowToBeUsed).value = nom
            .Range("D" & rowToBeUsed).value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).value = ObtenirNoGlIndicateur("Revenus de consultation")
            .Range("F" & rowToBeUsed).value = "Revenus de consultation"
            .Range("G" & rowToBeUsed).value = -CCur(ufEncRégularisation.txtHonoraires.value)
            .Range("I" & rowToBeUsed).value = desc
            .Range("J" & rowToBeUsed).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
             rowToBeUsed = rowToBeUsed + 1
       End If
        
        'Credit side - Frais divers
        If ufEncRégularisation.txtFraisDivers.value <> 0 Then
            .Range("A" & rowToBeUsed).value = nextJENo
            .Range("B" & rowToBeUsed).value = CDate(dt)
            .Range("C" & rowToBeUsed).value = nom
            .Range("D" & rowToBeUsed).value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).value = ObtenirNoGlIndicateur("Revenus frais de poste")
            .Range("F" & rowToBeUsed).value = "Revenus - Frais de poste"
            .Range("G" & rowToBeUsed).value = -CCur(ufEncRégularisation.txtFraisDivers.value)
            .Range("I" & rowToBeUsed).value = desc
            .Range("J" & rowToBeUsed).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
    
        'Credit side - TPS
        If ufEncRégularisation.txtTPS.value <> 0 Then
            .Range("A" & rowToBeUsed).value = nextJENo
            .Range("B" & rowToBeUsed).value = CDate(dt)
            .Range("C" & rowToBeUsed).value = nom
            .Range("D" & rowToBeUsed).value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).value = ObtenirNoGlIndicateur("TPS Facturée")
            .Range("F" & rowToBeUsed).value = "TPS percues"
            .Range("G" & rowToBeUsed).value = -CCur(ufEncRégularisation.txtTPS.value)
            .Range("I" & rowToBeUsed).value = desc
            .Range("J" & rowToBeUsed).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
    
        'Credit side - TVQ
        If ufEncRégularisation.txtTVQ.value <> 0 Then
            .Range("A" & rowToBeUsed).value = nextJENo
            .Range("B" & rowToBeUsed).value = CDate(dt)
            .Range("C" & rowToBeUsed).value = nom
            .Range("D" & rowToBeUsed).value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).value = ObtenirNoGlIndicateur("TVQ Facturée")
            .Range("F" & rowToBeUsed).value = "TVQ percues"
            .Range("G" & rowToBeUsed).value = -CCur(ufEncRégularisation.txtTVQ.value)
            .Range("I" & rowToBeUsed).value = desc
            .Range("J" & rowToBeUsed).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
        
        'Débit = Total de Honoraires, Frais Divers, TPS & TVQ
        Dim regulTotal As Currency
        regulTotal = CCur(ufEncRégularisation.txtHonoraires.value) + _
                    CCur(ufEncRégularisation.txtFraisDivers.value) + _
                    CCur(ufEncRégularisation.txtTPS.value) + _
                    CCur(ufEncRégularisation.txtTVQ.value)
    
        If regulTotal <> 0 Then
            .Range("A" & rowToBeUsed).value = nextJENo
            .Range("B" & rowToBeUsed).value = CDate(dt)
            .Range("C" & rowToBeUsed).value = nom
            .Range("D" & rowToBeUsed).value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).value = ObtenirNoGlIndicateur("Comptes Clients")
            .Range("F" & rowToBeUsed).value = "Comptes clients"
            .Range("H" & rowToBeUsed).value = -regulTotal
            .Range("I" & rowToBeUsed).value = desc
            .Range("J" & rowToBeUsed).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
    End With

    Application.ScreenUpdating = True

    Call Log_Record("modCC_Régularisation:REGUL_GL_Posting_Locally", "", startTime)

End Sub

