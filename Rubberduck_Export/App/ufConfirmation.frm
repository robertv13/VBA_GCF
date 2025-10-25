VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufConfirmation 
   Caption         =   "Confirmation des factures"
   ClientHeight    =   9120.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16560
   OleObjectBlob   =   "ufConfirmation.frx":0000
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
        .ColumnHeaders.Add , , vbNullString, 17
        .ColumnHeaders.Add , , " NoFact.", 57
        .ColumnHeaders.Add , , "    Date", 68
        .ColumnHeaders.Add , , "Nom du client", 424
        .ColumnHeaders.Add , , " Total Fact.", 80
        
        'Vérifiez que la collection Factures est définie et contient des données
        If Not Factures Is Nothing And Factures.count > 0 Then
            Dim nomClient As String
            Dim facture As Variant
            For Each facture In Factures
                If IsArray(facture) Then
                    Dim newItem As listItem
                    Set newItem = .ListItems.Add(, , vbNullString)
                    newItem.SubItems(1) = facture(0)
                    newItem.SubItems(2) = facture(1)
                    'Ajustement sur le nom du client
                    nomClient = Trim$(facture(2))
                    If Len(nomClient) > 60 Then
                        nomClient = Left$(nomClient, 60)
                    Else
                        nomClient = nomClient + Space(60 - Len(nomClient))
                    End If
'                    nomClient = Left$(nomClient, 55) & "   * Sélectionnée *"
                    newItem.SubItems(3) = nomClient
                    newItem.SubItems(4) = facture(3)
                Else
                    MsgBox "Erreur : L'élément n'est pas un tableau"
                End If
            Next facture
        Else
            Debug.Print "#091 - La collection Factures est vide ou non initialisée."
        End If
    End With
    
    ufConfirmation.shpConfirmation.Visible = False
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
    PDFInvoicePath = wsdADMIN.Range("PATH_DATA_FILES").Value & gFACT_PDF_PATH & _
                     Application.PathSeparator & noFacture & ".pdf"
                     
    'Open the invoice using Adobe Acrobat Reader
    If PDFInvoicePath <> vbNullString Then
        Dim strShell As String
        strShell = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe " & Chr$(34) & PDFInvoicePath & Chr$(34)
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
        ufConfirmation.txtTotalFacturesSélectionnées.Value = _
            Format$(ufConfirmation.txtTotalFacturesSélectionnées.Value + valeur, "###,##0.00 $")
        ufConfirmation.txtNbFacturesSélectionnées.Value = _
            ufConfirmation.txtNbFacturesSélectionnées.Value + 1
        If ufConfirmation.txtNbFacturesSélectionnées.Value > 0 Then
            ufConfirmation.shpConfirmation.Visible = True
            If ufConfirmation.txtNbFacturesSélectionnées.Value = 1 Then
                ufConfirmation.shpConfirmation.Caption = "Confirmer cette facture"
            Else
                ufConfirmation.shpConfirmation.Caption = "Confirmer les (" & _
                 ufConfirmation.txtNbFacturesSélectionnées.Value & ") factures sélectionnées"
            End If
        End If
    Else
        Call MarquerLigneSelectionnee(item)
        ufConfirmation.txtTotalFacturesSélectionnées = _
            Format$(ufConfirmation.txtTotalFacturesSélectionnées - valeur, "###,##0.00 $")
        ufConfirmation.txtNbFacturesSélectionnées = _
            ufConfirmation.txtNbFacturesSélectionnées - 1
        If ufConfirmation.txtNbFacturesSélectionnées.Value = 0 Then
            ufConfirmation.shpConfirmation.Visible = False
        Else
            If ufConfirmation.txtNbFacturesSélectionnées.Value = 1 Then
                ufConfirmation.shpConfirmation.Caption = "Confirmer cette facture"
            Else
                ufConfirmation.shpConfirmation.Caption = "Confirmer les (" & _
                 ufConfirmation.txtNbFacturesSélectionnées.Value & ") factures sélectionnées"
            End If
        End If
    End If

End Sub

'@Description ("Ajoute un petit message dans le tableau des factures à confirmer")
Public Sub MarquerLigneSelectionnee(item As listItem) '2025-06-17 @ 19:58

    'Vérifie si l'élément n'a pas déjà la mention "   - Sélectionnée -"
    If InStr(item.SubItems(3), "   - Sélectionnée -") = 0 Then
        item.SubItems(3) = Left$(item.SubItems(3), 60) & "   - Sélectionnée -"
    Else
        item.SubItems(3) = Left$(item.SubItems(3), 60)
    End If
    
End Sub

Private Sub shpCocherToutesCases_Click()

    Call modFAC_Confirmation.CocherToutesLesCases(ListView1)

End Sub

Private Sub shpDecocherToutesCases_Click()

    Call modFAC_Confirmation.DecocherToutesLesCases(ListView1)

End Sub

Private Sub shpConfirmation_Click()

    Call modFAC_Confirmation.ConfirmerSauvegardeConfirmationFacture

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    'Vérifie si le mode de fermeture est dû au clic sur le X du userForm (CloseMode = 0)
    If CloseMode = 0 And ufConfirmation.txtNbFacturesSélectionnées > 0 Then
        'Demande confirmation avant de fermer le UserForm
        Dim mess As String
        If ufConfirmation.txtNbFacturesSélectionnées.Value = 1 Then
            mess = ufConfirmation.txtNbFacturesSélectionnées.Value & " facture sélectionnée"
        Else
            mess = ufConfirmation.txtNbFacturesSélectionnées.Value & " factures sélectionnées"
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

    Call RetournerAuMenu
    
End Sub

Sub RetournerAuMenu()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufConfirmation:RetournerAuMenu", vbNullString, 0)

    ufConfirmation.Hide
    Unload ufConfirmation
    
    Call modDev_Utils.EnregistrerLogApplication("ufConfirmation:RetournerAuMenu", vbNullString, startTime)

    Call modAppli.QuitterFeuillePourMenu(wshMenuFAC, True)

End Sub
