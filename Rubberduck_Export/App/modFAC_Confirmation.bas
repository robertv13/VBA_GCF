Attribute VB_Name = "modFAC_Confirmation"
Option Explicit

Public invNo As String
Public Factures As Collection

Sub AfficherFormulaireConfirmation() '2025-03-12 @ 12:40

    'Aller chercher les factures à confirmer
    Call PreparerDonneesPourListView
    
    'Vérifiez si la collection de factures est vide
    If Factures Is Nothing Or Factures.count = 0 Then
        MsgBox "Il n'y a aucune facture à confirmer", vbInformation, "Toutes les factures sont déjà confirmées"
    Else
        'Charger le UserForm seulement avec une collection qui n'est pas vide
        Load ufConfirmation
        ufConfirmation.show vbModeless
        
        ufConfirmation.txtNbTotalFactures.Value = Format$(Factures.count, "#,##0")
        ufConfirmation.txtNbFacturesSélectionnées.Value = 0
        ufConfirmation.txtTotalFacturesSélectionnées.Value = Format$(0, "###,##0.00 $")
    End If

End Sub

Sub PreparerDonneesPourListView() '2025-03-12 @ 12:40

    Set Factures = New Collection
    
    Call ObtenirFactureAConfirmer("AC")
    
    Dim ws As Worksheet
    Set ws = wsdFAC_Entete
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AZ").End(xlUp).Row
    
    Dim invNo As String, dateFacture As String, nomClient As String, totalFacture As String
    Dim r As Long
    If lastUsedRow > 2 Then
        For r = 3 To lastUsedRow
            invNo = " " & ws.Range("AZ" & r).Value
            dateFacture = " " & Format$(ws.Range("BA" & r), wsdADMIN.Range("B1").Value)
            nomClient = ws.Range("BD" & r).Value
            totalFacture = Format$(ws.Range("BP" & r).Value, "###,##0.00 $")
            totalFacture = Space(13 - Len(totalFacture)) & totalFacture
            Factures.Add Array(invNo, dateFacture, nomClient, totalFacture)
        Next r
    End If

End Sub

Sub ObtenirFactureAConfirmer(AC_OR_C As String) '2025-03-12 @ 12:40

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Confirmation:ObtenirFactureAConfirmer", vbNullString, 0)
    
    'Utilisation de la feuille FAC_Entete
    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    
    'Utilisation du AF#2 dans wsdFAC_Entete
    
    'Effacer les données de la dernière utilisation
    ws.Range("AX6:AX10").ClearContents
    ws.Range("AX6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:nn:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_FAC_Entete[#All]")
    ws.Range("AX7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("AX2:AX3")
    ws.Range("AX3").Value = AC_OR_C
    ws.Range("AX8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("AZ1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("AZ2:BQ2")
    ws.Range("AX9").Value = rngResult.Address
        
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Qu'avons-nous comme résultat ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AZ").End(xlUp).Row
    ws.Range("AX10").Value = lastUsedRow - 2 & " lignes"
    
    If lastUsedRow > 3 Then
        With ws.Sort 'Sort - Inv_No
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("AZ3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Invoice Number
            .SetRange ws.Range("AZ3:BQ" & lastUsedRow) 'Set Range
            .Apply 'Apply Sort
         End With
     End If

    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Confirmation:ObtenirFactureAConfirmer", vbNullString, startTime)

End Sub

Sub CocherToutesLesCases(listView As listView) '2025-03-12 @ 12:40

    'On s'assure de commencer avec aucune ligne de sélectionnée
    ufConfirmation.txtNbFacturesSélectionnées.Value = 0
    ufConfirmation.txtTotalFacturesSélectionnées.Value = 0
    
    Dim valeur As Currency
    Dim i As Integer
    For i = 1 To listView.ListItems.count
        listView.ListItems(i).Checked = True
        Call MarquerLigneSelectionnee(listView.ListItems(i))
        valeur = CCur(Trim$(listView.ListItems(i).SubItems(4)))
        ufConfirmation.txtTotalFacturesSélectionnées.Value = _
            Format$(ufConfirmation.txtTotalFacturesSélectionnées.Value + valeur, "###,##0.00 $")
        ufConfirmation.txtNbFacturesSélectionnées.Value = _
            ufConfirmation.txtNbFacturesSélectionnées.Value + 1
    Next i
    
    If ufConfirmation.txtNbFacturesSélectionnées.Value = 1 Then
        ufConfirmation.shpConfirmation.Caption = "Confirmer cette facture"
    Else
        ufConfirmation.shpConfirmation.Caption = "Confirmer les (" & _
         ufConfirmation.txtNbFacturesSélectionnées.Value & ") factures sélectionnées"
    End If
    ufConfirmation.shpConfirmation.Visible = True
    
End Sub

Sub DecocherToutesLesCases(listView As listView) '2025-03-12 @ 12:40

    Dim i As Integer
    For i = 1 To listView.ListItems.count
        listView.ListItems(i).Checked = False
        Call MarquerLigneSelectionnee(listView.ListItems(i))
    Next i
    
    ufConfirmation.txtTotalFacturesSélectionnées = Format$(0, "###,##0.00 $")
    ufConfirmation.txtNbFacturesSélectionnées = 0
    ufConfirmation.shpConfirmation.Visible = False
    
End Sub

Public Sub MarquerLigneSelectionnee(item As listItem) '2025-03-12 @ 12:40

    'Vérifie si l'élément n'a pas déjà la mention "   - Sélectionnée -"
    If InStr(item.SubItems(3), "   - Sélectionnée -") = 0 Then
        item.SubItems(3) = Left$(item.SubItems(3), 60) & "   - Sélectionnée -"
    Else
        item.SubItems(3) = Left$(item.SubItems(3), 60)
    End If
    
End Sub

Public Sub ConfirmerSauvegardeConfirmationFacture()

    Dim uf As UserForm: Set uf = ufConfirmation
    
    If uf.ListView1.ListItems.count < 1 Then
        MsgBox "Vous n'avez sélectionné aucune facture à confirmer"
        Exit Sub
    Else
        Dim mess As String
        If uf.txtNbFacturesSélectionnées.Value = 1 Then
            mess = uf.txtNbFacturesSélectionnées.Value & " facture sélectionnée"
        Else
            mess = uf.txtNbFacturesSélectionnées.Value & " factures sélectionnées"
        End If
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Êtes-vous certain de vouloir procéder à la confirmation de" & _
                            vbNewLine & vbNewLine & "facture, avec " & mess & " ?", _
                            vbQuestion + vbYesNo, "Confirmation de traitement avec " & mess)
        If reponse = vbNo Then
            'Annule la confirmation si l'utilisateur répond Non
            GoTo exitSub
        End If
        Call MettreAJourConfirmationFacture
    End If
    
exitSub:
    'Libérer la mémoire
    Set uf = Nothing

End Sub

Sub MettreAJourConfirmationFacture() '2025-03-12 @ 12:40

    Dim uf As UserForm: Set uf = ufConfirmation
    
    Dim ligne As listItem
    
    uf.lblFactureEmConfirmation.Visible = True
    uf.txtNoFactureEnConfirmation.Visible = True

    Application.ScreenUpdating = True
    
    With uf.ListView1
        Dim i As Long
        'Parcourir chacune des lignes
        For i = 1 To .ListItems.count
            Set ligne = .ListItems(i)
            If ligne.Checked Then
                invNo = Trim$(ligne.SubItems(1))
                uf.txtNoFactureEnConfirmation.Value = invNo
                DoEvents
                Call MettreAJourStatutFacEnteteMaster(invNo)
                Call MettreAJourStatutFacEnteteLocale(invNo)
                DoEvents
                Call ComptabiliserConfirmationFacture(invNo)
                DoEvents
            End If
        Next i
    End With

    MsgBox "La confirmation des factures est complétée", vbOKOnly + vbInformation, "Confirmation de traitement"

    Unload ufConfirmation
    Call AfficherFormulaireConfirmation
    
    'Libérer la mémoire
    Set uf = Nothing
    
End Sub

Sub MettreAJourStatutFacEnteteMaster(invoice As String) '2025-03-12 @ 12:40

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Confirmation:MettreAJourStatutFacEnteteMaster", invoice, 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Entete$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    Dim sql As String
    'Open the recordset for the specified invoice
    sql = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & invoice & "'"
    recSet.Open sql, conn, 2, 3
    If Not recSet.EOF Then
        'Update AC_ouC with 'C'
        recSet.Fields(fFacEACouC - 1).Value = "C"
        recSet.Update
    Else
        'Handle the case where the specified invoice is not found
        MsgBox "La facture '" & invoice & "' n'existe pas!", vbCritical
    End If
    
    'Close recordset and connection
    recSet.Close
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Confirmation:MettreAJourStatutFacEnteteMaster", vbNullString, startTime)

End Sub

Sub MettreAJourStatutFacEnteteLocale(invoice As String) '2025-03-12 @ 12:40
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Confirmation:MettreAJourStatutFacEnteteLocale", invoice, 0)
    
    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=invoice, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim r As Long, rowToBeUpdated As Long, tecID As Long
    If Not foundRange Is Nothing Then
        r = foundRange.row
        ws.Cells(r, fFacEACouC).Value = "C"
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Entete."
    End If
    
    'Libérer la mémoire
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Confirmation:MettreAJourStatutFacEnteteLocale", vbNullString, startTime)

End Sub

Sub ComptabiliserConfirmationFacture(invoice As String) '2025-08-04 @ 07:19

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Confirmation:ComptabiliserConfirmationFacture", invoice, 0)

    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=invoice, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim r As Long
    If Not foundRange Is Nothing Then
        r = foundRange.row
        Dim dateFact As Date
        dateFact = Left$(ws.Cells(r, fFacEDateFacture).Value, 10)
        Dim hono As Currency
        hono = ws.Cells(r, fFacEHonoraires).Value
        Dim misc1 As Currency
        Dim misc2 As Currency
        Dim misc3 As Currency
        misc1 = ws.Cells(r, fFacEAutresFrais1).Value
        misc2 = ws.Cells(r, fFacEAutresFrais2).Value
        misc3 = ws.Cells(r, fFacEAutresFrais3).Value
        Dim tps As Currency
        Dim tvq As Currency
        tps = ws.Cells(r, fFacEMntTPS).Value
        tvq = ws.Cells(r, fFacEMntTVQ).Value
        
        'Déclaration et instanciation d'un objet GL_Entry
        Dim ecr As clsGL_Entry
        Set ecr = New clsGL_Entry
    
        'Remplissage des propriétés communes
        ecr.DateEcriture = dateFact
        ecr.description = ws.Cells(r, fFacENomClient).Value
        ecr.source = "FACTURE:" & invoice
        
        Dim codeGL As String
        Dim descGL As String
        
        'Débit Comptes Clients
        If hono + misc1 + misc2 + misc3 + tps + tvq <> 0 Then
            codeGL = Fn_NoCompteAPartirIndicateurCompte("Comptes Clients")
            descGL = Fn_DescriptionAPartirNoCompte(codeGL)
            ecr.AjouterLigne codeGL, descGL, hono + misc1 + misc2 + misc3 + tps + tvq, vbNullString
        End If
        
        'Honoraires
        If hono Then
            codeGL = Fn_NoCompteAPartirIndicateurCompte("Revenus de consultation")
            descGL = Fn_DescriptionAPartirNoCompte(codeGL)
            ecr.AjouterLigne codeGL, descGL, -hono, vbNullString
        End If
        
        'Miscellaneous Amount # 1 (misc1)
        If misc1 Then
            codeGL = Fn_NoCompteAPartirIndicateurCompte("Revenus frais de poste")
            descGL = Fn_DescriptionAPartirNoCompte(codeGL)
            ecr.AjouterLigne codeGL, descGL, -misc1, vbNullString
        End If
        
        'Miscellaneous Amount # 2 (misc2)
        If misc2 Then
            codeGL = Fn_NoCompteAPartirIndicateurCompte("Revenus sous-traitants")
            descGL = Fn_DescriptionAPartirNoCompte(codeGL)
            ecr.AjouterLigne codeGL, descGL, -misc2, vbNullString
        End If
        
        'Miscellaneous Amount # 3 (misc3)
        If misc3 Then
            codeGL = Fn_NoCompteAPartirIndicateurCompte("Revenus autres frais")
            descGL = Fn_DescriptionAPartirNoCompte(codeGL)
            ecr.AjouterLigne codeGL, descGL, -misc3, vbNullString
        End If
        
        'TPS à payer
        If tps Then
            codeGL = Fn_NoCompteAPartirIndicateurCompte("TPS Facturée")
            descGL = Fn_DescriptionAPartirNoCompte(codeGL)
            ecr.AjouterLigne codeGL, descGL, -tps, vbNullString
        End If
        
        'TVQ à payer
        If tvq Then
            codeGL = Fn_NoCompteAPartirIndicateurCompte("TVQ Facturée")
            descGL = Fn_DescriptionAPartirNoCompte(codeGL)
            ecr.AjouterLigne codeGL, descGL, -tvq, vbNullString
        End If
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Entete.", vbCritical
    End If
    
    'Écriture
    Call modGL_Stuff.AjouterEcritureGLADOPlusLocale(ecr, False)
    
    'Libérer la mémoire
    On Error Resume Next
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Confirmation:ComptabiliserConfirmationFacture", vbNullString, startTime)

End Sub


