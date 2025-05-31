VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufConfirmation 
   Caption         =   "Confirmation des factures"
   ClientHeight    =   9060.001
   ClientLeft      =   195
   ClientTop       =   780
   ClientWidth     =   16545
   OleObjectBlob   =   "ufConfirmation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ArgumentWithIncompatibleObjectType

Option Explicit

Private Sub UserForm_Initialize()
    
    With ListView1
        .View = lvwReport
        .CheckBoxes = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , , "", 17
        .ColumnHeaders.Add , , " NoFact.", 57
        .ColumnHeaders.Add , , "    Date", 68
        .ColumnHeaders.Add , , "Nom du client", 424
        .ColumnHeaders.Add , , " Total Fact.", 80
        
        'Vérifiez que la collection Factures est définie et contient des données
        If Not Factures Is Nothing And Factures.count > 0 Then
            Dim nomClient As String
            Dim Facture As Variant
            For Each Facture In Factures
                If IsArray(Facture) Then
                    Dim newItem As listItem
                    Set newItem = .ListItems.Add(, , "")
                    newItem.SubItems(1) = Facture(0)
                    newItem.SubItems(2) = Facture(1)
                    'Ajustement sur le nom du client
                    nomClient = Trim$(Facture(2))
                    If Len(nomClient) > 60 Then
                        nomClient = Left$(nomClient, 60)
                    Else
                        nomClient = nomClient + Space(60 - Len(nomClient))
                    End If
'                    nomClient = Left$(nomClient, 55) & "   * Sélectionnée *"
                    newItem.SubItems(3) = nomClient
                    newItem.SubItems(4) = Facture(3)
                Else
                    MsgBox "Erreur : L'élément n'est pas un tableau"
                End If
            Next Facture
        Else
            Debug.Print "#091 - La collection Factures est vide ou non initialisée."
        End If
    End With
    
    ufConfirmation.cmdConfirmation.Visible = False
    ufConfirmation.lblFactureEmConfirmation.Visible = False
    ufConfirmation.txtNoFactureEnConfirmation.Visible = False
    
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.listItem)

    'Met en srubrillance la ligne sélectionnée
    Set ListView1.SelectedItem = item
    
    'Accéder au contenu des colonnes de la ligne sélectionnée
    Dim noFacture As String
    Dim dateFacture As String
    Dim nomClient As String
    Dim totalFacture As String
    
    noFacture = Trim$(item.SubItems(1))
    
    Dim PDFInvoicePath As String
    PDFInvoicePath = wsdADMIN.Range("F5").value & FACT_PDF_PATH & _
                     Application.PathSeparator & noFacture & ".pdf"
                     
    'Open the invoice using Adobe Acrobat Reader
    If PDFInvoicePath <> "" Then
        Dim strShell As String
        strShell = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe " & Chr$(34) & PDFInvoicePath & Chr$(34)
        Debug.Print "#092 - " & strShell
        Shell strShell, vbNormalFocus
    Else
        MsgBox "Le fichier PDF de la facture n'existe pas : " & PDFInvoicePath, vbExclamation, "Fichier PDF manquant"
    End If
    
End Sub

Private Sub ListView1_ItemCheck(ByVal item As MSComctlLib.listItem)

    'Récupérer la valeur de la quatrième colonne
    Dim valeur As Currency
    valeur = CCur(Trim$(item.SubItems(4)))

    'Ajouter ou soustraire la valeur en fonction de l'état de la case à cocher
    If item.Checked Then
        Call MarquerLigneSelectionnee(item)
        ufConfirmation.txtTotalFacturesSélectionnées.value = _
            Format$(ufConfirmation.txtTotalFacturesSélectionnées.value + valeur, "###,##0.00 $")
        ufConfirmation.txtNbFacturesSélectionnées.value = _
            ufConfirmation.txtNbFacturesSélectionnées.value + 1
        If ufConfirmation.txtNbFacturesSélectionnées.value > 0 Then
            ufConfirmation.cmdConfirmation.Visible = True
            If ufConfirmation.txtNbFacturesSélectionnées.value = 1 Then
                ufConfirmation.cmdConfirmation.Caption = "Confirmer cette facture"
            Else
                ufConfirmation.cmdConfirmation.Caption = "Confirmer les (" & _
                 ufConfirmation.txtNbFacturesSélectionnées.value & ") factures sélectionnées"
            End If
        End If
    Else
        Call MarquerLigneSelectionnee(item)
        ufConfirmation.txtTotalFacturesSélectionnées = _
            Format$(ufConfirmation.txtTotalFacturesSélectionnées - valeur, "###,##0.00 $")
        ufConfirmation.txtNbFacturesSélectionnées = _
            ufConfirmation.txtNbFacturesSélectionnées - 1
        If ufConfirmation.txtNbFacturesSélectionnées.value = 0 Then
            ufConfirmation.cmdConfirmation.Visible = False
        Else
            If ufConfirmation.txtNbFacturesSélectionnées.value = 1 Then
                ufConfirmation.cmdConfirmation.Caption = "Confirmer cette facture"
            Else
                ufConfirmation.cmdConfirmation.Caption = "Confirmer les (" & _
                 ufConfirmation.txtNbFacturesSélectionnées.value & ") factures sélectionnées"
            End If
        End If
    End If

End Sub

Public Sub MarquerLigneSelectionnee(item As listItem)

    'Vérifie si l'élément n'a pas déjà la mention "   - Sélectionnée -"
    If InStr(item.SubItems(3), "   - Sélectionnée -") = 0 Then
        item.SubItems(3) = Left$(item.SubItems(3), 60) & "   - Sélectionnée -"
    Else
        item.SubItems(3) = Left$(item.SubItems(3), 60)
    End If
    
End Sub

Private Sub cmdCocherToutesCases_Click()

    Call CocherToutesLesCases(ListView1)

End Sub

Private Sub cmdDécocherToutesCases_Click()

    Call DecocherToutesLesCases(ListView1)

End Sub

Private Sub cmdConfirmation_Click()

    If ListView1.ListItems.count < 1 Then
        MsgBox "Vous n'avez sélectionné aucune facture à confirmer"
        Exit Sub
    Else
        Dim mess As String
        If ufConfirmation.txtNbFacturesSélectionnées.value = 1 Then
            mess = ufConfirmation.txtNbFacturesSélectionnées.value & " facture sélectionnée"
        Else
            mess = ufConfirmation.txtNbFacturesSélectionnées.value & " factures sélectionnées"
        End If
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Êtes-vous certain de vouloir procéder à la confirmation de" & _
                            vbNewLine & vbNewLine & "facture, avec " & mess & " ?", _
                            vbQuestion + vbYesNo, "Confirmation de traitement avec " & mess)
        If reponse = vbNo Then
            'Annule la confirmation si l'utilisateur répond Non
            Exit Sub
        End If
        Call Confirmation_Mise_À_Jour
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    'Vérifie si le mode de fermeture est dû au clic sur le X du userForm (CloseMode = 0)
    If CloseMode = 0 And ufConfirmation.txtNbFacturesSélectionnées > 0 Then
        'Demande confirmation avant de fermer le UserForm
        Dim mess As String
        If ufConfirmation.txtNbFacturesSélectionnées.value = 1 Then
            mess = ufConfirmation.txtNbFacturesSélectionnées.value & " facture sélectionnée"
        Else
            mess = ufConfirmation.txtNbFacturesSélectionnées.value & " factures sélectionnées"
        End If
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Êtes-vous certain de vouloir quitter la confirmation de" & _
                            vbNewLine & vbNewLine & "facture, avec " & mess & " ?", _
                            vbQuestion + vbYesNo, "Confirmation de fermeture avec " & mess)
        If reponse = vbNo Then
            'Annule la fermeture si l'utilisateur répond Non
            Cancel = True
        End If
    End If
    
End Sub

Private Sub UserForm_Terminate()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufConfirmation:UserForm_Terminate", "", 0)

    ufConfirmation.Hide
    Unload ufConfirmation
    
    If ufConfirmation.Name = "ufConfirmation" Then
        On Error GoTo MenuSelect
        wshMenuFAC.Select
        On Error GoTo 0
    Else
        wshMenu.Select
    End If
    
    GoTo Exit_Sub
    
MenuSelect:
    wshMenu.Activate
    wshMenu.Select
    
Exit_Sub:
    Call Log_Record("ufConfirmation:UserForm_Terminate", "", startTime)

End Sub

