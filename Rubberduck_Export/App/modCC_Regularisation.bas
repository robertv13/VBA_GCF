Attribute VB_Name = "modCC_Regularisation"
'@Folder("Saisie_Encaissement")

Option Explicit

'Variables globales pour le module
Public regulNo As Long
Public gNextJENo As Long

Sub MAJ_Regularisation() '2025-01-14 @ 12:00
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Regularisation:MAJ_Regularisation", "", 0)
    
    With wshENC_Saisie
        'As-t-on les champs obligatoires ?
        If .Range("F5").Value = Empty Or _
           .Range("K5").Value = Empty Or _
           .Range("F7").Value = Empty Or _
           .Range("K7").Value = 0 Then
            MsgBox "Assurez-vous d'avoir..." & vbNewLine & vbNewLine & _
                "1. Un client valide" & vbNewLine & _
                "2. Une date de régularisation" & vbNewLine & _
                "3. Un type de transaction et" & vbNewLine & _
                "4. Le montant de la régularisation" & vbNewLine & vbNewLine & _
                "AVANT de tenter de sauvegarder la régularisation.", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Le montant de la régularisation doit être appliqué intégralement
        If .Range("K9").Value <> CCur(ufEncRégularisation.txtTotalFacture) Then
            MsgBox "Assurez-vous que le montant de la régularisation soit ÉGAL" & vbNewLine & _
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
        dateRegul = wshENC_Saisie.Range("K5").Value
        nomClient = wshENC_Saisie.Range("F5").Value
        descRegul = wshENC_Saisie.Range("F9").Value

        Call REGUL_GL_Posting_DB(regulNo, dateRegul, nomClient, descRegul)
        Call REGUL_GL_Posting_Locally(regulNo, dateRegul, nomClient, descRegul)
        
        MsgBox "La régularisation '" & regulNo & "' a été enregistré avec succès", vbOKOnly + vbInformation, "Confirmation de traitement"
        
        'Fermer le UserForm
        Unload ufEncRégularisation
    
        Call Régularisation_Add_New 'Reset the form
        
        Call AjusteLibelleEncaissement("Banque")
    
        .Range("K5").Value = Format$(Date, wsdADMIN.Range("B1").Value)
        .Range("B1").Select

        'De retour à la saisie du client
        .Range("F5").Select
        
    End With
    
Clean_Exit:

    Call Log_Record("modCC_Regularisation:MAJ_Regularisation", "", startTime)

End Sub

Sub Régularisation_Add_New() '2024-08-21 @ 14:58

    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Regularisation:Régularisation_Add_New", "", 0)

    Call ENC_Clear_Cells
    
    Call Log_Record("modCC_Regularisation:Régularisation_Add_New", "", startTime)
    
End Sub

Sub REGUL_Add_DB() 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Regularisation:REGUL_Add_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
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
    If IsNull(rs.Fields("MaxRegulNo").Value) Then
        lr = 0
    Else
        lr = rs.Fields("MaxRegulNo").Value
    End If
    
    'Calculate the new PmtNo
    regulNo = lr + 1

    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
        rs.Fields(fREGULRegulID - 1).Value = regulNo
        rs.Fields(fREGULInvNo - 1).Value = ufEncRégularisation.cbbNoFacture
        rs.Fields(fREGULDate - 1).Value = CDate(wshENC_Saisie.Range("K5").Value)
        rs.Fields(fREGULClientID - 1).Value = wshENC_Saisie.clientCode
        rs.Fields(fREGULClientNom - 1).Value = wshENC_Saisie.Range("F5").Value
        rs.Fields(fREGULHono - 1).Value = CCur(ufEncRégularisation.txtHonoraires)
        rs.Fields(fREGULFrais - 1).Value = CCur(ufEncRégularisation.txtFraisDivers)
        rs.Fields(fREGULTPS - 1).Value = CCur(ufEncRégularisation.txtTPS)
        rs.Fields(fREGULTVQ - 1).Value = CCur(ufEncRégularisation.txtTVQ)
        rs.Fields(fREGULDescription - 1).Value = wshENC_Saisie.Range("F9").Value
        rs.Fields(fREGULTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    'Update the recordset (create the record)
    rs.Update
    
    'Close recordset and connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modCC_Regularisation:REGUL_Add_DB", "", startTime)
    
End Sub

Sub REGUL_Add_Locally() '2024-08-22 @ 10:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Regularisation:REGUL_Add_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentRegulNo As Long
    currentRegulNo = regulNo
    
    'What is the last used row in CC_Régularisations ?
    Dim ws As Worksheet: Set ws = wsdCC_Regularisations
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    ws.Range("A" & rowToBeUsed).Value = regulNo
    ws.Range("B" & rowToBeUsed).Value = ufEncRégularisation.cbbNoFacture
    ws.Range("C" & rowToBeUsed).Value = CDate(wshENC_Saisie.Range("K5").Value)
    ws.Range("D" & rowToBeUsed).Value = wshENC_Saisie.clientCode
    ws.Range("E" & rowToBeUsed).Value = wshENC_Saisie.Range("F5").Value
    ws.Range("F" & rowToBeUsed).Value = CCur(ufEncRégularisation.txtHonoraires)
    ws.Range("G" & rowToBeUsed).Value = CCur(ufEncRégularisation.txtFraisDivers)
    ws.Range("H" & rowToBeUsed).Value = CCur(ufEncRégularisation.txtTPS)
    ws.Range("I" & rowToBeUsed).Value = CCur(ufEncRégularisation.txtTVQ)
    ws.Range("J" & rowToBeUsed).Value = wshENC_Saisie.Range("F9").Value
    ws.Range("K" & rowToBeUsed).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    
    Application.ScreenUpdating = True

    Call Log_Record("modCC_Regularisation:REGUL_Add_Locally", "", startTime)

End Sub

Sub REGUL_Update_Comptes_Clients_DB() 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Regularisation:REGUL_Update_Comptes_Clients_DB", "", 0)
    
    Dim errMsg As String
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Comptes_Clients$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Open the recordset for the specified invoice
    Dim Inv_No As String
    Inv_No = ufEncRégularisation.cbbNoFacture.Value
    
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & Inv_No & "'"
    rs.Open strSQL, conn, 2, 3
    If Not (rs.BOF Or rs.EOF) Then
        Dim mntRegulTotal As Double
        mntRegulTotal = CDbl(ufEncRégularisation.txtHonoraires.Value) + _
                        CDbl(ufEncRégularisation.txtFraisDivers.Value) + _
                        CDbl(ufEncRégularisation.txtTPS.Value) + _
                        CDbl(ufEncRégularisation.txtTVQ.Value)

        'Mettre à jour Régularisation totale
        rs.Fields(fFacCCTotalRegul - 1).Value = rs.Fields(fFacCCTotalRegul - 1).Value + mntRegulTotal
        'Mettre à jour le solde de la facture
        rs.Fields(fFacCCBalance - 1).Value = rs.Fields(fFacCCBalance - 1).Value + mntRegulTotal
        'Mettre à jour Status
        If rs.Fields(fFacCCBalance - 1).Value = 0 Then
            rs.Fields(fFacCCStatus - 1).Value = "Paid"
        Else
            rs.Fields(fFacCCStatus - 1).Value = "Unpaid"
        End If
        rs.Update
    Else
        MsgBox "Problème avec la facture '" & Inv_No & "'" & vbNewLine & vbNewLine & _
               "Contactez le développeur SVP", vbCritical, "Impossible de trouver la facture dans Comptes_Clients"
    End If
    'Update the recordset (create the record)
    rs.Close
    
    'Close recordset and connection
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modCC_Regularisation:REGUL_Update_Comptes_Clients_DB", "", startTime)
    
End Sub

Sub REGUL_Update_Comptes_Clients_Locally() '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Regularisation:REGUL_Update_Comptes_Clients_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'Set the range to look for
    Dim ws As Worksheet: Set ws = wsdFAC_Comptes_Clients
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim Inv_No As String
    Inv_No = ufEncRégularisation.cbbNoFacture.Value
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=Inv_No, LookIn:=xlValues, LookAt:=xlWhole)

    Dim rowToBeUpdated As Long
    If Not foundRange Is Nothing Then
        rowToBeUpdated = foundRange.row
        ws.Cells(rowToBeUpdated, fFacCCTotalRegul).Value = ws.Cells(rowToBeUpdated, fFacCCTotalRegul).Value + _
                                                            CCur(ufEncRégularisation.txtHonoraires.Value) + _
                                                            CCur(ufEncRégularisation.txtFraisDivers.Value) + _
                                                            CCur(ufEncRégularisation.txtTPS.Value) + _
                                                            CCur(ufEncRégularisation.txtTVQ.Value)
        ws.Cells(rowToBeUpdated, fFacCCBalance).Value = ws.Cells(rowToBeUpdated, fFacCCBalance).Value + _
                                                            CCur(ufEncRégularisation.txtHonoraires.Value) + _
                                                            CCur(ufEncRégularisation.txtFraisDivers.Value) + _
                                                            CCur(ufEncRégularisation.txtTPS.Value) + _
                                                            CCur(ufEncRégularisation.txtTVQ.Value)
 
        'Est-ce que le solde de la facture est à 0,00 $ ?
        If ws.Cells(rowToBeUpdated, fFacCCBalance).Value = 0 Then
            ws.Cells(rowToBeUpdated, fFacCCStatus) = "Paid"
        Else
            ws.Cells(rowToBeUpdated, fFacCCStatus) = "Unpaid"
        End If
    Else
        MsgBox "La facture '" & Inv_No & "' n'existe pas dans FAC_Comptes_Clients.", vbCritical
    End If
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modCC_Regularisation:REGUL_Update_Comptes_Clients_Locally", "", startTime)

End Sub

Sub REGUL_GL_Posting_DB(no As Long, dt As Date, nom As String, desc As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Regularisation:REGUL_GL_Posting_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
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
    If IsNull(rs.Fields("MaxEJNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").Value
    End If
    
    'Calculate the new ID
    gNextJENo = lastJE + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Crédit - Honoraires
    If ufEncRégularisation.txtHonoraires.Value <> 0 Then
        rs.AddNew
            'Cnstruction des champs
            rs.Fields(fGlTNoEntrée - 1).Value = gNextJENo
            rs.Fields(fGlTDate - 1).Value = Format$(dt, "yyyy-mm-dd")
            rs.Fields(fGlTDescription - 1).Value = nom
            rs.Fields(fGlTSource - 1).Value = "RÉGULARISATION:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).Value = ObtenirNoGlIndicateur("Revenus de consultation")
            rs.Fields(fGlTCompte - 1).Value = "Revenus de consultation"
            rs.Fields(fGlTDébit - 1).Value = -CCur(ufEncRégularisation.txtHonoraires.Value)
            rs.Fields(fGlTAutreRemarque - 1).Value = desc
            rs.Fields(fGlTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        rs.Update
    End If
    
    'Crédit - Frais Divers
    If ufEncRégularisation.txtFraisDivers.Value <> 0 Then
        rs.AddNew
            'Cnstruction des champs
            rs.Fields(fGlTNoEntrée - 1).Value = gNextJENo
            rs.Fields(fGlTDate - 1).Value = Format$(dt, "yyyy-mm-dd")
            rs.Fields(fGlTDescription - 1).Value = nom
            rs.Fields(fGlTSource - 1).Value = "RÉGULARISATION:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).Value = ObtenirNoGlIndicateur("Revenus frais de poste")
            rs.Fields(fGlTCompte - 1).Value = "Revenus - Frais de poste"
            rs.Fields(fGlTDébit - 1).Value = -CCur(ufEncRégularisation.txtFraisDivers.Value)
            rs.Fields(fGlTAutreRemarque - 1).Value = desc
            rs.Fields(fGlTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        rs.Update
    End If
    
    'Crédit - TPS
    If ufEncRégularisation.txtTPS.Value <> 0 Then
        rs.AddNew
            'Cnstruction des champs
            rs.Fields(fGlTNoEntrée - 1).Value = gNextJENo
            rs.Fields(fGlTDate - 1).Value = Format$(dt, "yyyy-mm-dd")
            rs.Fields(fGlTDescription - 1).Value = nom
            rs.Fields(fGlTSource - 1).Value = "RÉGULARISATION:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).Value = ObtenirNoGlIndicateur("TPS Facturée")
            rs.Fields(fGlTCompte - 1).Value = "TPS percues"
            rs.Fields(fGlTDébit - 1).Value = -CCur(ufEncRégularisation.txtTPS.Value)
            rs.Fields(fGlTAutreRemarque - 1).Value = desc
            rs.Fields(fGlTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        rs.Update
    End If
    
    'Crédit - TVQ
    If ufEncRégularisation.txtTVQ.Value <> 0 Then
        rs.AddNew
            'Cnstruction des champs
            rs.Fields(fGlTNoEntrée - 1).Value = gNextJENo
            rs.Fields(fGlTDate - 1).Value = Format$(dt, "yyyy-mm-dd")
            rs.Fields(fGlTDescription - 1).Value = nom
            rs.Fields(fGlTSource - 1).Value = "RÉGULARISATION:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).Value = ObtenirNoGlIndicateur("TVQ Facturée")
            rs.Fields(fGlTCompte - 1).Value = "TVQ percues"
            rs.Fields(fGlTDébit - 1).Value = -CCur(ufEncRégularisation.txtTVQ.Value)
            rs.Fields(fGlTAutreRemarque - 1).Value = desc
            rs.Fields(fGlTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        rs.Update
    End If
    
    'Débit = Total de Honoraires, Frais Divers, TPS & TVQ
    Dim regulTotal As Currency
    regulTotal = CCur(ufEncRégularisation.txtHonoraires.Value) + _
                    CCur(ufEncRégularisation.txtFraisDivers.Value) + _
                    CCur(ufEncRégularisation.txtTPS.Value) + _
                    CCur(ufEncRégularisation.txtTVQ.Value)
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields(fGlTNoEntrée - 1).Value = gNextJENo
        rs.Fields(fGlTDate - 1).Value = Format$(dt, "yyyy-mm-dd")
        rs.Fields(fGlTDescription - 1).Value = nom
        rs.Fields(fGlTSource - 1).Value = "RÉGULARISATION:" & Format$(no, "00000")
        rs.Fields(fGlTNoCompte - 1).Value = ObtenirNoGlIndicateur("Comptes Clients")
        rs.Fields(fGlTCompte - 1).Value = "Comptes clients"
        rs.Fields(fGlTCrédit - 1).Value = -regulTotal
        rs.Fields(fGlTAutreRemarque - 1).Value = desc
        rs.Fields(fGlTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
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
    
    Call Log_Record("modCC_Regularisation:REGUL_GL_Posting_DB", "", startTime)

End Sub

Sub REGUL_GL_Posting_Locally(no As Long, dt As Date, nom As String, desc As String)  'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCC_Regularisation:REGUL_GL_Posting_Locally", "", 0)

    Application.ScreenUpdating = False

    'What is the last used row in GL_Trans ?
    Dim ws As Worksheet: Set ws = wsdGL_Trans
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1

    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    With ws
        'Credit side - Honoraires
        If ufEncRégularisation.txtHonoraires.Value <> 0 Then
            .Range("A" & rowToBeUsed).Value = gNextJENo
            .Range("B" & rowToBeUsed).Value = CDate(dt)
            .Range("C" & rowToBeUsed).Value = nom
            .Range("D" & rowToBeUsed).Value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).Value = ObtenirNoGlIndicateur("Revenus de consultation")
            .Range("F" & rowToBeUsed).Value = "Revenus de consultation"
            .Range("G" & rowToBeUsed).Value = -CCur(ufEncRégularisation.txtHonoraires.Value)
            .Range("I" & rowToBeUsed).Value = desc
            .Range("J" & rowToBeUsed).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
             rowToBeUsed = rowToBeUsed + 1
       End If
        
        'Credit side - Frais divers
        If ufEncRégularisation.txtFraisDivers.Value <> 0 Then
            .Range("A" & rowToBeUsed).Value = gNextJENo
            .Range("B" & rowToBeUsed).Value = CDate(dt)
            .Range("C" & rowToBeUsed).Value = nom
            .Range("D" & rowToBeUsed).Value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).Value = ObtenirNoGlIndicateur("Revenus frais de poste")
            .Range("F" & rowToBeUsed).Value = "Revenus - Frais de poste"
            .Range("G" & rowToBeUsed).Value = -CCur(ufEncRégularisation.txtFraisDivers.Value)
            .Range("I" & rowToBeUsed).Value = desc
            .Range("J" & rowToBeUsed).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
    
        'Credit side - TPS
        If ufEncRégularisation.txtTPS.Value <> 0 Then
            .Range("A" & rowToBeUsed).Value = gNextJENo
            .Range("B" & rowToBeUsed).Value = CDate(dt)
            .Range("C" & rowToBeUsed).Value = nom
            .Range("D" & rowToBeUsed).Value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).Value = ObtenirNoGlIndicateur("TPS Facturée")
            .Range("F" & rowToBeUsed).Value = "TPS percues"
            .Range("G" & rowToBeUsed).Value = -CCur(ufEncRégularisation.txtTPS.Value)
            .Range("I" & rowToBeUsed).Value = desc
            .Range("J" & rowToBeUsed).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
    
        'Credit side - TVQ
        If ufEncRégularisation.txtTVQ.Value <> 0 Then
            .Range("A" & rowToBeUsed).Value = gNextJENo
            .Range("B" & rowToBeUsed).Value = CDate(dt)
            .Range("C" & rowToBeUsed).Value = nom
            .Range("D" & rowToBeUsed).Value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).Value = ObtenirNoGlIndicateur("TVQ Facturée")
            .Range("F" & rowToBeUsed).Value = "TVQ percues"
            .Range("G" & rowToBeUsed).Value = -CCur(ufEncRégularisation.txtTVQ.Value)
            .Range("I" & rowToBeUsed).Value = desc
            .Range("J" & rowToBeUsed).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
        
        'Débit = Total de Honoraires, Frais Divers, TPS & TVQ
        Dim regulTotal As Currency
        regulTotal = CCur(ufEncRégularisation.txtHonoraires.Value) + _
                    CCur(ufEncRégularisation.txtFraisDivers.Value) + _
                    CCur(ufEncRégularisation.txtTPS.Value) + _
                    CCur(ufEncRégularisation.txtTVQ.Value)
    
        If regulTotal <> 0 Then
            .Range("A" & rowToBeUsed).Value = gNextJENo
            .Range("B" & rowToBeUsed).Value = CDate(dt)
            .Range("C" & rowToBeUsed).Value = nom
            .Range("D" & rowToBeUsed).Value = "RÉGULARISATION:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).Value = ObtenirNoGlIndicateur("Comptes Clients")
            .Range("F" & rowToBeUsed).Value = "Comptes clients"
            .Range("H" & rowToBeUsed).Value = -regulTotal
            .Range("I" & rowToBeUsed).Value = desc
            .Range("J" & rowToBeUsed).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
    End With

    Application.ScreenUpdating = True

    Call Log_Record("modCC_Regularisation:REGUL_GL_Posting_Locally", "", startTime)

End Sub

