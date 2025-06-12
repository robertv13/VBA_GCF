VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufEncR�gularisation 
   Caption         =   "R�gularisation des Comptes Clients"
   ClientHeight    =   5736
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10080
   OleObjectBlob   =   "ufEncR�gularisation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufEncR�gularisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Set ws = wshENC_Saisie
    
    Dim factureRange As Range
    Set factureRange = ws.Range("F12:F36")
    
    'Vider le ComboBox avant de charger de nouvelles donn�es
    ufEncR�gularisation.cbbNoFacture.Clear

    'Parcourir la plage et charger les factures
    Dim row As Range
    For Each row In factureRange.Rows
        If row.Cells(1, 1).value <> "" Then
            ufEncR�gularisation.cbbNoFacture.AddItem row.Cells(1, 1).value
        End If
    Next row

    Call EffaceDonn�esR�gularisation
    
    'V�rifier les �l�ments dans le ComboBox
    If Me.cbbNoFacture.ListCount >= 1 Then
        Me.cbbNoFacture.ListIndex = 0 'S�lectionne automatiquement le premier (et unique) �l�ment
        Call cbbNoFacture_AfterUpdate ' Appelle explicitement l'�v�nement AfterUpdate
    Else
        MsgBox "Aucune facture, avec solde, n'existe pour ce client.", vbExclamation
        Exit Sub
    End If
    
End Sub

Private Sub cbbNoFacture_Change()

    Call cbbNoFacture_AfterUpdate
    
End Sub

Private Sub cbbNoFacture_AfterUpdate()

    Dim wsCC As Worksheet
    Set wsCC = wsdFAC_Comptes_Clients
    
    Dim invNo As String
    invNo = ufEncR�gularisation.cbbNoFacture.value
    
    Dim rngTrouve As Range
    Set rngTrouve = wsCC.Columns(fFacCCInvNo).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)

    Dim soldeFacture As Currency
    Dim dateFacture As String
    
    If Not rngTrouve Is Nothing Then
        'Si la valeur est trouv�e, r�cup�rer d'autres colonnes
        soldeFacture = CCur(rngTrouve.offset(0, 10).value)
        dateFacture = Format$(rngTrouve.offset(0, 1).value, wsdADMIN.Range("B1").value)
        ufEncR�gularisation.lblDateFactureData.Caption = dateFacture
        ufEncR�gularisation.lblTotalFactureValue.Caption = FormatCurrency(soldeFacture, 2)
        
        ufEncR�gularisation.txtTotalFacture.value = ""
        ufEncR�gularisation.txtHonoraires.value = ""
        ufEncR�gularisation.txtFraisDivers.value = ""
        ufEncR�gularisation.txtTPS.value = ""
        ufEncR�gularisation.txtTVQ.value = ""
        
        ufEncR�gularisation.lblTotalFactureAjuste.Caption = Format$(soldeFacture, "###,##0.00 $")
    Else
        'Si la valeur n'est pas trouv�e
        MsgBox "La facture " & invNo & " n'a pas �t� trouv�e.", vbExclamation
    End If
    
    ufEncR�gularisation.cmbAccepte.Visible = False
    ufEncR�gularisation.cmbRejete.Visible = False
    
End Sub

Private Sub txtTotalFacture_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim totalFacture As Currency
    ufEncR�gularisation.txtTotalFacture.Text = Replace(ufEncR�gularisation.txtTotalFacture.Text, ".", ",")
    
    If ufEncR�gularisation.txtTotalFacture.Text <> "" And IsNumeric(ufEncR�gularisation.txtTotalFacture.value) = True Then
        totalFacture = CCur(ufEncR�gularisation.txtTotalFacture.Text)
        Debug.Print "#093 - totalFacture = " & totalFacture
    
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
        
        ufEncR�gularisation.txtTotalFacture.value = Format$(totalFacture, "###,##0.00 $")
        ufEncR�gularisation.lblTotalFactureAjuste.Caption = Format$(lblTotalFactureValue + totalFacture, "###,##0.00 $")
        ufEncR�gularisation.txtHonoraires.value = Format$(honoraires, "###,##0.00 $")
        ufEncR�gularisation.txtFraisDivers.value = Format$(fraisDivers, "###,##0.00 $")
        ufEncR�gularisation.txtTPS.value = Format$(tps, "###,##0.00 $")
        ufEncR�gularisation.txtTVQ.value = Format$(tvq, "###,##0.00 $")
        
        ufEncR�gularisation.cmbAccepte.Visible = True
        ufEncR�gularisation.cmbRejete.Visible = True
    End If
    
End Sub

Private Sub txtHonoraires_AfterUpdate()

    ufEncR�gularisation.txtHonoraires.Text = Replace(ufEncR�gularisation.txtHonoraires.Text, ".", ",")
    ufEncR�gularisation.txtHonoraires.Text = Format$(ufEncR�gularisation.txtHonoraires.Text, "###,##0.00 $")
    Call VerifieMontantsSaisis
    
End Sub

Private Sub txtFraisDivers_AfterUpdate()

    ufEncR�gularisation.txtFraisDivers.Text = Replace(ufEncR�gularisation.txtFraisDivers.Text, ".", ",")
    ufEncR�gularisation.txtFraisDivers.Text = Format$(ufEncR�gularisation.txtFraisDivers.Text, "###,##0.00 $")
    Call VerifieMontantsSaisis

End Sub

Private Sub txtTPS_AfterUpdate()

    ufEncR�gularisation.txtTPS.Text = Replace(ufEncR�gularisation.txtTPS.Text, ".", ",")
    ufEncR�gularisation.txtTPS.Text = Format$(ufEncR�gularisation.txtTPS.Text, "###,##0.00 $")
    Call VerifieMontantsSaisis

End Sub

Private Sub txtTVQ_AfterUpdate()

    ufEncR�gularisation.txtTVQ.Text = Replace(ufEncR�gularisation.txtTVQ.Text, ".", ",")
    ufEncR�gularisation.txtTVQ.Text = Format$(ufEncR�gularisation.txtTVQ.Text, "###,##0.00 $")
    Call VerifieMontantsSaisis

End Sub

Private Sub VerifieMontantsSaisis()

    If ufEncR�gularisation.txtTotalFacture.Text <> "" Then
        With ufEncR�gularisation
            Debug.Print "#095 - " & CCur(.txtTotalFacture.Text) & " <> ? " & CCur(.txtHonoraires.Text) & "+" & CCur(.txtFraisDivers.Text) & "+" & CCur(.txtTPS.Text) & "+" & CCur(.txtTVQ.Text)
            If CCur(.txtTotalFacture.Text) = CCur(.txtHonoraires.Text) + _
                                        CCur(.txtFraisDivers.Text) + _
                                        CCur(.txtTPS.Text) + _
                                        CCur(.txtTVQ.Text) Then
                .txtTotalFacture.ForeColor = vbBlack
                ufEncR�gularisation.cmbAccepte.Visible = True
            Else
                .txtTotalFacture.ForeColor = vbRed
                ufEncR�gularisation.cmbAccepte.Visible = False
            End If
        End With
    End If
    
End Sub

Private Sub cmbRejete_Click()

    Call EffaceDonn�esR�gularisation
    cmbAccepte.Visible = False
    cmbRejete.Visible = False
    txtTotalFacture.SetFocus
    
End Sub

Private Sub cmbAccepte_Click()

    Dim reponse As VbMsgBoxResult

    'Afficher une bo�te de message avec les boutons Oui et Non
    reponse = MsgBox("Toujours pr�t � continuer le traitement ?", vbYesNo + vbQuestion, "Confirmation avant traitement")
    
    'V�rifie la r�ponse de l'utilisateur
    If reponse = vbYes Then
        Call MAJ_Regularisation
    Else
        Exit Sub
    End If
    
End Sub

Sub EffaceDonn�esR�gularisation()

    With ufEncR�gularisation
        
        'Montants de la r�gularisation
        .txtTotalFacture.value = ""
        .txtHonoraires.value = ""
        .txtFraisDivers.value = ""
        .txtTPS.value = ""
        .txtTVQ.value = ""
        
    End With
    
End Sub
