VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufEncRégularisation 
   Caption         =   "Régularisation des Comptes Clients"
   ClientHeight    =   5730
   ClientLeft      =   -60
   ClientTop       =   -240
   ClientWidth     =   10020
   OleObjectBlob   =   "ufEncRégularisation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufEncRégularisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private wrappers As Collection

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Set ws = wshENC_Saisie
    
    Dim factureRange As Range
    Set factureRange = ws.Range("F12:F36")
    
    'Vider le ComboBox avant de charger de nouvelles données
    ufEncRégularisation.cmbNoFacture.Clear

    'Parcourir la plage et charger les factures
    Dim row As Range
    For Each row In factureRange.Rows
        If row.Cells(1, 1).Value <> vbNullString Then
            ufEncRégularisation.cmbNoFacture.AddItem row.Cells(1, 1).Value
        End If
    Next row

    Call EffacerDonneesRegularisation
    
    'Vérifier les éléments dans le ComboBox
    If Me.cmbNoFacture.ListCount >= 1 Then
        Me.cmbNoFacture.ListIndex = 0 'Sélectionne automatiquement le premier (et unique) élément
        Call cmbNoFacture_AfterUpdate ' Appelle explicitement l'événement AfterUpdate
    Else
        MsgBox "Aucune facture, avec solde, n'existe pour ce client.", vbExclamation
        Exit Sub
    End If
    
    Call InitialiserSurveillanceForm(Me, wrappers)
    
End Sub

Private Sub cmbNoFacture_Change()

    Call cmbNoFacture_AfterUpdate
    
End Sub

Private Sub cmbNoFacture_AfterUpdate()

    Dim wsCC As Worksheet
    Set wsCC = wsdFAC_Comptes_Clients
    
    Dim invNo As String
    invNo = ufEncRégularisation.cmbNoFacture.Value
    
    Dim rngTrouve As Range
    Set rngTrouve = wsCC.Columns(fFacCCInvNo).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)

    Dim soldeFacture As Currency
    Dim dateFacture As String
    
    If Not rngTrouve Is Nothing Then
        'Si la valeur est trouvée, récupérer d'autres colonnes
        soldeFacture = CCur(rngTrouve.offset(0, 10).Value)
        dateFacture = Format$(rngTrouve.offset(0, 1).Value, wsdADMIN.Range("USER_DATE_FORMAT").Value)
        ufEncRégularisation.lblDateFactureData.caption = dateFacture
        ufEncRégularisation.lblTotalFactureValue.caption = FormatCurrency(soldeFacture, 2)
        
        ufEncRégularisation.txtTotalFacture.Value = vbNullString
        ufEncRégularisation.txtHonoraires.Value = vbNullString
        ufEncRégularisation.txtFraisDivers.Value = vbNullString
        ufEncRégularisation.txtTPS.Value = vbNullString
        ufEncRégularisation.txtTVQ.Value = vbNullString
        
        ufEncRégularisation.lblTotalFactureAjuste.caption = Format$(soldeFacture, "###,##0.00 $")
    Else
        'Si la valeur n'est pas trouvée
        MsgBox "La facture " & invNo & " n'a pas été trouvée.", vbExclamation
    End If
    
    ufEncRégularisation.shpAccepte.Visible = False
    ufEncRégularisation.shpRejete.Visible = False
    
End Sub

Private Sub txtTotalFacture_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim totalFacture As Currency
    ufEncRégularisation.txtTotalFacture.text = Replace(ufEncRégularisation.txtTotalFacture.text, ".", ",")
    
    If ufEncRégularisation.txtTotalFacture.text <> vbNullString And IsNumeric(ufEncRégularisation.txtTotalFacture.Value) = True Then
        totalFacture = CCur(ufEncRégularisation.txtTotalFacture.text)
    
        Dim honoraires As Currency, fraisDivers As Currency
        Dim tps As Currency, tvq As Currency
    
        Dim gstRate As Double, pstRate As Double
        gstRate = Fn_Get_Tax_Rate(Date, "TPS")
        pstRate = Fn_Get_Tax_Rate(Date, "TVQ")
        
        tps = Round(totalFacture / (1 + gstRate + pstRate) * gstRate, 2)
        tvq = Round(totalFacture / (1 + gstRate + pstRate) * pstRate, 2)
        
        fraisDivers = 0
        honoraires = totalFacture - fraisDivers - tps - tvq
        Dim ecartArrondissement As Currency
        ecartArrondissement = totalFacture - honoraires - fraisDivers - tps - tvq
        If ecartArrondissement <> 0 Then
            Debug.Print "#094 - " & ecartArrondissement
            honoraires = honoraires + ecartArrondissement
        End If
        
        ufEncRégularisation.txtTotalFacture.Value = Format$(totalFacture, "###,##0.00 $")
        ufEncRégularisation.lblTotalFactureAjuste.caption = Format$(lblTotalFactureValue + totalFacture, "###,##0.00 $")
        ufEncRégularisation.txtHonoraires.Value = Format$(honoraires, "###,##0.00 $")
        ufEncRégularisation.txtFraisDivers.Value = Format$(fraisDivers, "###,##0.00 $")
        ufEncRégularisation.txtTPS.Value = Format$(tps, "###,##0.00 $")
        ufEncRégularisation.txtTVQ.Value = Format$(tvq, "###,##0.00 $")
        
        ufEncRégularisation.shpAccepte.Visible = True
        ufEncRégularisation.shpRejete.Visible = True
    End If
    
End Sub

Private Sub txtHonoraires_AfterUpdate()

    ufEncRégularisation.txtHonoraires.text = Replace(ufEncRégularisation.txtHonoraires.text, ".", ",")
    ufEncRégularisation.txtHonoraires.text = Format$(ufEncRégularisation.txtHonoraires.text, "###,##0.00 $")
    Call VerifierMontantsSaisis
    
End Sub

Private Sub txtFraisDivers_AfterUpdate()

    ufEncRégularisation.txtFraisDivers.text = Replace(ufEncRégularisation.txtFraisDivers.text, ".", ",")
    ufEncRégularisation.txtFraisDivers.text = Format$(ufEncRégularisation.txtFraisDivers.text, "###,##0.00 $")
    Call VerifierMontantsSaisis

End Sub

Private Sub txtTPS_AfterUpdate()

    ufEncRégularisation.txtTPS.text = Replace(ufEncRégularisation.txtTPS.text, ".", ",")
    ufEncRégularisation.txtTPS.text = Format$(ufEncRégularisation.txtTPS.text, "###,##0.00 $")
    Call VerifierMontantsSaisis

End Sub

Private Sub txtTVQ_AfterUpdate()

    ufEncRégularisation.txtTVQ.text = Replace(ufEncRégularisation.txtTVQ.text, ".", ",")
    ufEncRégularisation.txtTVQ.text = Format$(ufEncRégularisation.txtTVQ.text, "###,##0.00 $")
    Call VerifierMontantsSaisis

End Sub

Private Sub VerifierMontantsSaisis()

    If ufEncRégularisation.txtTotalFacture.text <> vbNullString Then
        With ufEncRégularisation
            If CCur(.txtTotalFacture.text) = CCur(.txtHonoraires.text) + _
                                        CCur(.txtFraisDivers.text) + _
                                        CCur(.txtTPS.text) + _
                                        CCur(.txtTVQ.text) Then
                .txtTotalFacture.ForeColor = vbBlack
                ufEncRégularisation.shpAccepte.Visible = True
            Else
                .txtTotalFacture.ForeColor = vbRed
                ufEncRégularisation.shpAccepte.Visible = False
            End If
        End With
    End If
    
End Sub

Private Sub shpRejete_Click()

    Call EffacerDonneesRegularisation
    shpAccepte.Visible = False
    shpRejete.Visible = False
    txtTotalFacture.SetFocus
    
End Sub

Private Sub shpAccepte_Click()

    Dim reponse As VbMsgBoxResult

    'Afficher une boîte de message avec les boutons Oui et Non
    reponse = MsgBox("Toujours prêt à continuer le traitement ?", vbYesNo + vbQuestion, "Confirmation avant traitement")
    
    'Vérifie la réponse de l'utilisateur
    If reponse = vbYes Then
        Call SauvegarderRegularisation
    Else
        Exit Sub
    End If
    
End Sub

Sub EffacerDonneesRegularisation()

    With ufEncRégularisation
        
        'Montants de la régularisation
        .txtTotalFacture.Value = vbNullString
        .txtHonoraires.Value = vbNullString
        .txtFraisDivers.Value = vbNullString
        .txtTPS.Value = vbNullString
        .txtTVQ.Value = vbNullString
        
    End With
    
End Sub

