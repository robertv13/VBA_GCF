Attribute VB_Name = "modFAC_Confirmation"
Option Explicit

Public invNo As String
Public Factures As Collection

Sub Afficher_ufConfirmation() '2025-03-12 @ 12:40

    'Aller chercher les factures à confirmer
    Call PrepareDonneesPourListView
    
    'Vérifiez si la collection de factures est vide
    If Factures Is Nothing Then
        MsgBox "Il n'y a aucune facture à confirmer", vbInformation, "Toutes les factures ont été confirmées"
    ElseIf Factures.count = 0 Then
        MsgBox "Il n'y a aucune facture à confirmer", vbInformation, "Toutes les factures ont été confirmées"
    Else
        'Charger le UserForm seulement avec une collection qui n'est pas vide
        Load ufConfirmation
        ufConfirmation.show vbModeless
        
        ufConfirmation.txtNbTotalFactures.Value = Format$(Factures.count, "#,##0")
        ufConfirmation.txtNbFacturesSélectionnées.Value = 0
        ufConfirmation.txtTotalFacturesSélectionnées.Value = Format$(0, "###,##0.00 $")
    End If

End Sub

Sub PrepareDonneesPourListView() '2025-03-12 @ 12:40

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

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirFactureAConfirmer", "", 0)
    
    'Utilisation de la feuille FAC_Entête
    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    
    'Utilisation du AF#2 dans wsdFAC_Entete
    
    'Effacer les données de la dernière utilisation
    ws.Range("AX6:AX10").ClearContents
    ws.Range("AX6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_FAC_Entête[#All]")
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

    Call Log_Record("modFAC_Confirmation:ObtenirFactureAConfirmer", "", startTime)

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
        ufConfirmation.cmdConfirmation.Caption = "Confirmer cette facture"
    Else
        ufConfirmation.cmdConfirmation.Caption = "Confirmer les (" & _
         ufConfirmation.txtNbFacturesSélectionnées.Value & ") factures sélectionnées"
    End If
    ufConfirmation.cmdConfirmation.Visible = True
    
End Sub

Sub DecocherToutesLesCases(listView As listView) '2025-03-12 @ 12:40

    Dim i As Integer
    For i = 1 To listView.ListItems.count
        listView.ListItems(i).Checked = False
        Call MarquerLigneSelectionnee(listView.ListItems(i))
    Next i
    
    ufConfirmation.txtTotalFacturesSélectionnées = Format$(0, "###,##0.00 $")
    ufConfirmation.txtNbFacturesSélectionnées = 0
    ufConfirmation.cmdConfirmation.Visible = False
    
End Sub

Public Sub MarquerLigneSelectionnee(item As listItem) '2025-03-12 @ 12:40

    'Vérifie si l'élément n'a pas déjà la mention "   - Sélectionnée -"
    If InStr(item.SubItems(3), "   - Sélectionnée -") = 0 Then
        item.SubItems(3) = Left$(item.SubItems(3), 60) & "   - Sélectionnée -"
    Else
        item.SubItems(3) = Left$(item.SubItems(3), 60)
    End If
    
End Sub

Sub Confirmation_Mise_À_Jour() '2025-03-12 @ 12:40

    Dim ligne As listItem
    
    ufConfirmation.lblFactureEmConfirmation.Visible = True
    ufConfirmation.txtNoFactureEnConfirmation.Visible = True

    Application.ScreenUpdating = True
    
    With ufConfirmation.ListView1
        Dim i As Long
        'Parcourir chacune des lignes
        For i = 1 To .ListItems.count
            Set ligne = .ListItems(i)
            If ligne.Checked Then
                invNo = Trim$(ligne.SubItems(1))
                ufConfirmation.txtNoFactureEnConfirmation.Value = invNo
                DoEvents
                Call MAJ_Statut_Facture_Entête_BD_MASTER(invNo)
                Call MAJ_Statut_Facture_Entête_Local(invNo)
                DoEvents
                Call Construire_GL_Posting_Confirmation(invNo)
                DoEvents
            End If
        Next i
    End With

    MsgBox "La confirmation des factures est complétée", vbOKOnly + vbInformation, "Confirmation de traitement"

    Unload ufConfirmation
    Call Afficher_ufConfirmation
    
End Sub

Sub MAJ_Statut_Facture_Entête_BD_MASTER(invoice As String) '2025-03-12 @ 12:40

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Entête_BD_MASTER", invoice, 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Entête$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
              ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim sql As String
    'Open the recordset for the specified invoice
    sql = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & invoice & "'"
    rs.Open sql, conn, 2, 3
    If Not rs.EOF Then
        'Update AC_ouC with 'C'
        rs.Fields(fFacEACouC - 1).Value = "C"
        rs.Update
    Else
        'Handle the case where the specified invoice is not found
        MsgBox "La facture '" & invoice & "' n'existe pas!", vbCritical
    End If
    
    'Close recordset and connection
    rs.Close
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Entête_BD_MASTER", "", startTime)

End Sub

Sub MAJ_Statut_Facture_Entête_Local(invoice As String) '2025-03-12 @ 12:40
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Entête_Local", invoice, 0)
    
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
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Entête."
    End If
    
    'Libérer la mémoire
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Entête_Local", "", startTime)

End Sub

Sub Construire_GL_Posting_Confirmation(invoice As String) '2025-03-12 @ 12:42

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:Construire_GL_Posting_Confirmation", invoice, 0)

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
        Dim misc1 As Currency, misc2 As Currency, misc3 As Currency
        misc1 = ws.Cells(r, fFacEAutresFrais1).Value
        misc2 = ws.Cells(r, fFacEAutresFrais2).Value
        misc3 = ws.Cells(r, fFacEAutresFrais3).Value
        Dim tps As Currency, tvq As Currency
        tps = ws.Cells(r, fFacEMntTPS).Value
        tvq = ws.Cells(r, fFacEMntTVQ).Value
        
        Dim descGL_Trans As String, Source As String
        descGL_Trans = ws.Cells(r, fFacENomClient).Value
        Source = "FACTURE:" & invoice
        
        Dim MyArray(1 To 7, 1 To 4) As String
        
        'AR amount
        If hono + misc1 + misc2 + misc3 + tps + tvq Then
            MyArray(1, 1) = ObtenirNoGlIndicateur("Comptes Clients")
            MyArray(1, 2) = "Comptes clients"
            MyArray(1, 3) = hono + misc1 + misc2 + misc3 + tps + tvq
            MyArray(1, 4) = ""
        End If
        
        'Professional Fees (hono)
        If hono Then
            MyArray(2, 1) = ObtenirNoGlIndicateur("Revenus de consultation")
            MyArray(2, 2) = "Revenus de consultation"
            MyArray(2, 3) = -hono
            MyArray(2, 4) = ""
        End If
        
        'Miscellaneous Amount # 1 (misc1)
        If misc1 Then
            MyArray(3, 1) = ObtenirNoGlIndicateur("Revenus frais de poste")
            MyArray(3, 2) = "Revenus - Frais de poste"
            MyArray(3, 3) = -misc1
            MyArray(3, 4) = ""
        End If
        
        'Miscellaneous Amount # 2 (misc2)
        If misc2 Then
            MyArray(4, 1) = ObtenirNoGlIndicateur("Revenus sous-traitants")
            MyArray(4, 2) = "Revenus - Sous-traitants"
            MyArray(4, 3) = -misc2
            MyArray(4, 4) = ""
        End If
        
        'Miscellaneous Amount # 3 (misc3)
        If misc3 Then
            MyArray(5, 1) = ObtenirNoGlIndicateur("Revenus autres frais")
            MyArray(5, 2) = "Revenus - Autres Frais"
            MyArray(5, 3) = -misc3
            MyArray(5, 4) = ""
        End If
        
        'GST to pay (tps)
        If tps Then
            MyArray(6, 1) = ObtenirNoGlIndicateur("TPS Facturée")
            MyArray(6, 2) = "TPS percues"
            MyArray(6, 3) = -tps
            MyArray(6, 4) = ""
        End If
        
        'PST to pay (tvq)
        If tvq Then
            MyArray(7, 1) = ObtenirNoGlIndicateur("TVQ Facturée")
            MyArray(7, 2) = "TVQ percues"
            MyArray(7, 3) = -tvq
            MyArray(7, 4) = ""
        End If
        
        'Mise à jour du posting GL des confirmations de facture
        Dim GLEntryNo As Long
        Call GL_Posting_To_DB(dateFact, descGL_Trans, Source, MyArray, GLEntryNo)
        Call GL_Posting_Locally(dateFact, descGL_Trans, Source, MyArray, GLEntryNo)
        
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Entête.", vbCritical
    End If
    
    'Libérer la mémoire
    On Error Resume Next
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Call Log_Record("modFAC_Confirmation:Construire_GL_Posting_Confirmation", "", startTime)

End Sub


