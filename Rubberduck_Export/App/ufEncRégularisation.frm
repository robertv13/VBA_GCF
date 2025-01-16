VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufEncR�gularisation 
   Caption         =   "R�gularisation des Comptes Clients"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   465
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
        If row.Cells(1, 1).Value <> "" Then
            ufEncR�gularisation.cbbNoFacture.AddItem row.Cells(1, 1).Value
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
    Set wsCC = wshFAC_Comptes_Clients
    
    Dim invNo As String
    invNo = ufEncR�gularisation.cbbNoFacture.Value
    
    Dim rngTrouve As Range
    Set rngTrouve = wsCC.Columns(fFacCCInvNo).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)

    Dim soldeFacture As Currency
    Dim dateFacture As String
    
    If Not rngTrouve Is Nothing Then
        'Si la valeur est trouv�e, r�cup�rer d'autres colonnes
        soldeFacture = CCur(rngTrouve.offset(0, 10).Value)
        dateFacture = Format$(rngTrouve.offset(0, 1).Value, wshAdmin.Range("B1").Value)
        ufEncR�gularisation.lblDateFactureData.Caption = dateFacture
        ufEncR�gularisation.lblTotalFactureValue.Caption = FormatCurrency(soldeFacture, 2)
        
        ufEncR�gularisation.txtTotalFacture.Value = ""
        ufEncR�gularisation.txtHonoraires.Value = ""
        ufEncR�gularisation.txtFraisDivers.Value = ""
        ufEncR�gularisation.txtTPS.Value = ""
        ufEncR�gularisation.txtTVQ.Value = ""
        
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
'    Debug.Print "1 - '" & ufEncR�gularisation.txtTotalFacture.Text & "' - '" & ufEncR�gularisation.txtTotalFacture.Value & "'"
    ufEncR�gularisation.txtTotalFacture.Text = Replace(ufEncR�gularisation.txtTotalFacture.Text, ".", ",")
'    Debug.Print "2 - '" & ufEncR�gularisation.txtTotalFacture.Text & "' - '" & ufEncR�gularisation.txtTotalFacture.Value & "'"
    Debug.Print "Text  ? " & IsNumeric(ufEncR�gularisation.txtTotalFacture.Text)
    Debug.Print "Value ? " & IsNumeric(ufEncR�gularisation.txtTotalFacture.Value)
    
    If ufEncR�gularisation.txtTotalFacture.Text <> "" And IsNumeric(ufEncR�gularisation.txtTotalFacture.Value) = True Then
        totalFacture = CCur(ufEncR�gularisation.txtTotalFacture.Text)
        Debug.Print "totalFacture = " & totalFacture
    
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
            Debug.Print ecartArrondissement
            honoraires = honoraires + ecartArrondissement
        End If
        
        ufEncR�gularisation.txtTotalFacture.Value = Format$(totalFacture, "###,##0.00 $")
        ufEncR�gularisation.lblTotalFactureAjuste.Caption = Format$(lblTotalFactureValue + totalFacture, "###,##0.00 $")
        ufEncR�gularisation.txtHonoraires.Value = Format$(honoraires, "###,##0.00 $")
        ufEncR�gularisation.txtFraisDivers.Value = Format$(fraisDivers, "###,##0.00 $")
        ufEncR�gularisation.txtTPS.Value = Format$(tps, "###,##0.00 $")
        ufEncR�gularisation.txtTVQ.Value = Format$(tvq, "###,##0.00 $")
        
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
            Debug.Print CCur(.txtTotalFacture.Text) & " <> ? " & CCur(.txtHonoraires.Text) & "+" & CCur(.txtFraisDivers.Text) & "+" & CCur(.txtTPS.Text) & "+" & CCur(.txtTVQ.Text)
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
'        'Solde actuel
'        .lblTotalFactureValue.Caption = ""
        
        'Montants de la r�gularisation
        .txtTotalFacture.Value = ""
        .txtHonoraires.Value = ""
        .txtFraisDivers.Value = ""
        .txtTPS.Value = ""
        .txtTVQ.Value = ""
        
'        'Nouveau solde
'        .lblTotalFactureAjuste.Caption = ""
    End With
    
End Sub
