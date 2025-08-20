Attribute VB_Name = "modCC_Regularisation"
'@Folder("Saisie_Encaissement")

Option Explicit

'Variables globales pour le module
Public regulNo As Long
Public gNextJENo As Long

Sub SauvegarderRegularisation() '2025-01-14 @ 12:00
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:SauvegarderRegularisation", vbNullString, 0)
    
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
        Call AjouterRegularisationBDMaster
        Call AjouterRegularisationBDLocale
        
        'Update FAC_Comptes_Clients
        Call MettreAJourRegulComptesClientsBDMaster
        Call MettreAJourRegulComptesClientsBDLocale
                
        'Prepare G/L posting
        Dim noRegul As String, nomClient As String, descRegul As String
        Dim dateRegul As Date
        Dim montantRegul As Currency
        dateRegul = wshENC_Saisie.Range("K5").Value
        nomClient = wshENC_Saisie.Range("F5").Value
        descRegul = wshENC_Saisie.Range("F9").Value
        
        Call ComptabiliserRegularisation(regulNo, dateRegul, nomClient, descRegul)

        MsgBox "La régularisation '" & regulNo & "' a été enregistré avec succès", vbOKOnly + vbInformation, "Confirmation de traitement"
        
        'Fermer le UserForm
        Unload ufEncRégularisation
    
        Call PreparerNouvelleRegularisation 'Reset the form
        
        Call AjusterLibelleDansEncaissement("Banque")
    
        .Range("K5").Value = Format$(Date, wsdADMIN.Range("B1").Value)
        .Range("B1").Select

        'De retour à la saisie du client
        .Range("F5").Select
        
    End With
    
Clean_Exit:

    Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:SauvegarderRegularisation", vbNullString, startTime)

End Sub

Sub PreparerNouvelleRegularisation() '2024-08-21 @ 14:58

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:PreparerNouvelleRegularisation", vbNullString, 0)

    Call NettoyerFeuilleEncaissement
    
    Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:PreparerNouvelleRegularisation", vbNullString, startTime)
    
End Sub

Sub AjouterRegularisationBDMaster() 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:AjouterRegularisationBDMaster", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "CC_Regularisations$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxRegulNo As Long
    strSQL = "SELECT MAX(RegulID) AS MaxRegulNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxPmtNo
    recSet.Open strSQL, conn
    
    'Get the last used row
    Dim lr As Long
    If IsNull(recSet.Fields("MaxRegulNo").Value) Then
        lr = 0
    Else
        lr = recSet.Fields("MaxRegulNo").Value
    End If
    
    'Calculate the new PmtNo
    regulNo = lr + 1

    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    recSet.Close
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    recSet.AddNew
        recSet.Fields(fREGULRegulID - 1).Value = regulNo
        recSet.Fields(fREGULInvNo - 1).Value = ufEncRégularisation.cmbNoFacture
        recSet.Fields(fREGULDate - 1).Value = CDate(wshENC_Saisie.Range("K5").Value)
        recSet.Fields(fREGULClientID - 1).Value = wshENC_Saisie.clientCode
        recSet.Fields(fREGULClientNom - 1).Value = wshENC_Saisie.Range("F5").Value
        recSet.Fields(fREGULHono - 1).Value = CCur(ufEncRégularisation.txtHonoraires)
        recSet.Fields(fREGULFrais - 1).Value = CCur(ufEncRégularisation.txtFraisDivers)
        recSet.Fields(fREGULTPS - 1).Value = CCur(ufEncRégularisation.txtTPS)
        recSet.Fields(fREGULTVQ - 1).Value = CCur(ufEncRégularisation.txtTVQ)
        recSet.Fields(fREGULDescription - 1).Value = wshENC_Saisie.Range("F9").Value
        recSet.Fields(fREGULTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    'Update the recordset (create the record)
    recSet.Update
    
    'Close recordset and connection
    recSet.Close
    Set recSet = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:AjouterRegularisationBDMaster", vbNullString, startTime)
    
End Sub

Sub AjouterRegularisationBDLocale() '2024-08-22 @ 10:38
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:AjouterRegularisationBDLocale", vbNullString, 0)
    
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
    ws.Range("B" & rowToBeUsed).Value = ufEncRégularisation.cmbNoFacture
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

    Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:AjouterRegularisationBDLocale", vbNullString, startTime)

End Sub

Sub MettreAJourRegulComptesClientsBDMaster() 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:MettreAJourRegulComptesClientsBDMaster", vbNullString, 0)
    
    Dim errMsg As String
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Comptes_Clients$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'Open the recordset for the specified invoice
    Dim Inv_No As String
    Inv_No = ufEncRégularisation.cmbNoFacture.Value
    
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & Inv_No & "'"
    recSet.Open strSQL, conn, 2, 3
    If Not (recSet.BOF Or recSet.EOF) Then
        Dim mntRegulTotal As Double
        mntRegulTotal = CDbl(ufEncRégularisation.txtHonoraires.Value) + _
                        CDbl(ufEncRégularisation.txtFraisDivers.Value) + _
                        CDbl(ufEncRégularisation.txtTPS.Value) + _
                        CDbl(ufEncRégularisation.txtTVQ.Value)

        'Mettre à jour Régularisation totale
        recSet.Fields(fFacCCTotalRegul - 1).Value = recSet.Fields(fFacCCTotalRegul - 1).Value + mntRegulTotal
        'Mettre à jour le solde de la facture
        recSet.Fields(fFacCCBalance - 1).Value = recSet.Fields(fFacCCBalance - 1).Value + mntRegulTotal
        'Mettre à jour Status
        If recSet.Fields(fFacCCBalance - 1).Value = 0 Then
            recSet.Fields(fFacCCStatus - 1).Value = "Paid"
        Else
            recSet.Fields(fFacCCStatus - 1).Value = "Unpaid"
        End If
        recSet.Update
    Else
        MsgBox "Problème avec la facture '" & Inv_No & "'" & vbNewLine & vbNewLine & _
               "Contactez le développeur SVP", vbCritical, "Impossible de trouver la facture dans Comptes_Clients"
    End If
    'Update the recordset (create the record)
    recSet.Close
    
    'Close recordset and connection
    Set recSet = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:MettreAJourRegulComptesClientsBDMaster", vbNullString, startTime)
    
End Sub

Sub MettreAJourRegulComptesClientsBDLocale() '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:MettreAJourRegulComptesClientsBDLocale", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    'Set the range to look for
    Dim ws As Worksheet: Set ws = wsdFAC_Comptes_Clients
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim Inv_No As String
    Inv_No = ufEncRégularisation.cmbNoFacture.Value
    
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
    
    Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:MettreAJourRegulComptesClientsBDLocale", vbNullString, startTime)

End Sub

Sub ComptabiliserRegularisation(no As Long, dt As Date, nom As String, desc As String) '2025-07-24 @ 07:02
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:ComptabiliserRegularisation", vbNullString, 0)
    
    Dim ws As Worksheet
    Set ws = wshTEC_Evaluation
    
    Dim uf As UserForm
    Set uf = ufEncRégularisation
    
    Dim honoraires As Currency
    Dim fraisDivers As Currency
    Dim tps As Currency
    Dim tvq As Currency
    Dim comptesClients As Currency
    
    Dim glComptesClients As String, descGLComptesClients As String
    Dim glHonoraires As String, descGLHonoraires As String
    Dim glFraisDivers As String, descGLFraisDivers As String
    Dim glTPS As String, descGLTPS As String
    Dim glTVQ As String, descGLTVQ As String
    
    honoraires = CCur(uf.txtHonoraires.Value)
    fraisDivers = CCur(uf.txtFraisDivers.Value)
    tps = CCur(uf.txtTPS.Value)
    tvq = CCur(uf.txtTVQ.Value)
    comptesClients = honoraires + fraisDivers + tps + tvq
    
    'Comptes de GL et description du poste
    glComptesClients = Fn_NoCompteAPartirIndicateurCompte("Comptes Clients")
    descGLComptesClients = Fn_DescriptionAPartirNoCompte(glComptesClients)
    glHonoraires = Fn_NoCompteAPartirIndicateurCompte("Revenus de consultation")
    descGLHonoraires = Fn_DescriptionAPartirNoCompte(glHonoraires)
    glFraisDivers = Fn_NoCompteAPartirIndicateurCompte("Revenus frais de poste")
    descGLFraisDivers = Fn_DescriptionAPartirNoCompte(glFraisDivers)
    glTPS = Fn_NoCompteAPartirIndicateurCompte("TPS Facturée")
    descGLTPS = Fn_DescriptionAPartirNoCompte(glTPS)
    glTVQ = Fn_NoCompteAPartirIndicateurCompte("TVQ Facturée")
    descGLTVQ = Fn_DescriptionAPartirNoCompte(glTVQ)
    
    'Déclaration et instanciation d'un objet GL_Entry
    Dim ecr As clsGL_Entry
    Set ecr = New clsGL_Entry

    'Remplissage des propriétés communes
    ecr.DateEcriture = dt
    ecr.description = nom
    ecr.source = "RÉGULARISATION:" & Format$(no, "00000")

    'Ajoute autant de lignes que nécessaire
    If honoraires <> 0 Then
        ecr.AjouterLigne glHonoraires, descGLHonoraires, -honoraires, ""
    End If
    
    If fraisDivers <> 0 Then
        ecr.AjouterLigne glFraisDivers, descGLFraisDivers, -fraisDivers, ""
    End If

    If tps <> 0 Then
        ecr.AjouterLigne glTPS, descGLTPS, -tps, ""
    End If
    
    If tvq <> 0 Then
        ecr.AjouterLigne glTVQ, descGLTVQ, -tvq, ""
    End If
    
    If comptesClients <> 0 Then
        ecr.AjouterLigne glComptesClients, descGLComptesClients, comptesClients, ""
    End If
    
    '--- Écriture ---
    Call modGL_Stuff.AjouterEcritureGLADOPlusLocale(ecr, False)
    
    MsgBox "L'écriture pour cette régularisation s'est complétée" & vbNewLine & vbNewLine & _
            "avec succès", _
            vbInformation + vbOKOnly, _
            "Comptabilisation de l'écriture de régularisation"
    
    Call modDev_Utils.EnregistrerLogApplication("modCC_Regularisation:ComptabiliserRegularisation", vbNullString, startTime)
            
End Sub

