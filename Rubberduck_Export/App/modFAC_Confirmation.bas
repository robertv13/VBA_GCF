Attribute VB_Name = "modFAC_Confirmation"
Option Explicit

Public invNo As String
Public Factures As Collection

Sub Afficher_ufConfirmation() '2025-03-12 @ 12:40

    'Aller chercher les factures � confirmer
    Call PrepareDonneesPourListView
    
    'V�rifiez si la collection de factures est vide
    If Factures Is Nothing Then
        MsgBox "Il n'y a aucune facture � confirmer", vbInformation, "Toutes les factures ont �t� confirm�es"
    ElseIf Factures.count = 0 Then
        MsgBox "Il n'y a aucune facture � confirmer", vbInformation, "Toutes les factures ont �t� confirm�es"
    Else
        'Charger le UserForm seulement avec une collection qui n'est pas vide
        Load ufConfirmation
        ufConfirmation.show vbModeless
        
        ufConfirmation.txtNbTotalFactures.value = Format$(Factures.count, "#,##0")
        ufConfirmation.txtNbFacturesS�lectionn�es.value = 0
        ufConfirmation.txtTotalFacturesS�lectionn�es.value = Format$(0, "###,##0.00 $")
    End If

End Sub

Sub PrepareDonneesPourListView() '2025-03-12 @ 12:40

    Set Factures = New Collection
    
    Call ObtenirFactureAConfirmer("AC")
    
    Dim ws As Worksheet
    Set ws = wsdFAC_Ent�te
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AZ").End(xlUp).row
    
    Dim invNo As String, dateFacture As String, nomClient As String, totalFacture As String
    Dim r As Long
    If lastUsedRow > 2 Then
        For r = 3 To lastUsedRow
            invNo = " " & ws.Range("AZ" & r).value
            dateFacture = " " & Format$(ws.Range("BA" & r), wsdADMIN.Range("B1").value)
            nomClient = ws.Range("BD" & r).value
            totalFacture = Format$(ws.Range("BP" & r).value, "###,##0.00 $")
            totalFacture = Space(13 - Len(totalFacture)) & totalFacture
            Factures.Add Array(invNo, dateFacture, nomClient, totalFacture)
        Next r
    End If

End Sub

Sub ObtenirFactureAConfirmer(AC_OR_C As String) '2025-03-12 @ 12:40

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirFactureAConfirmer", "", 0)
    
    'Utilisation de la feuille FAC_Ent�te
    Dim ws As Worksheet: Set ws = wsdFAC_Ent�te
    
    'Utilisation du AF#2 dans wsdFAC_Ent�te
    
    'Effacer les donn�es de la derni�re utilisation
    ws.Range("AX6:AX10").ClearContents
    ws.Range("AX6").value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'D�finir le range pour la source des donn�es en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_FAC_Ent�te[#All]")
    ws.Range("AX7").value = rngData.Address
    
    'D�finir le range des crit�res
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("AX2:AX3")
    ws.Range("AX3").value = AC_OR_C
    ws.Range("AX8").value = rngCriteria.Address
    
    'D�finir le range des r�sultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("AZ1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("AZ2:BQ2")
    ws.Range("AX9").value = rngResult.Address
        
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Qu'avons-nous comme r�sultat ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AZ").End(xlUp).row
    ws.Range("AX10").value = lastUsedRow - 2 & " lignes"
    
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

    'Lib�rer la m�moire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing

    Call Log_Record("modFAC_Confirmation:ObtenirFactureAConfirmer", "", startTime)

End Sub

Sub CocherToutesLesCases(listView As listView) '2025-03-12 @ 12:40

    'On s'assure de commencer avec aucune ligne de s�lectionn�e
    ufConfirmation.txtNbFacturesS�lectionn�es.value = 0
    ufConfirmation.txtTotalFacturesS�lectionn�es.value = 0
    
    Dim valeur As Currency
    Dim i As Integer
    For i = 1 To listView.ListItems.count
        listView.ListItems(i).Checked = True
        Call MarquerLigneSelectionnee(listView.ListItems(i))
        valeur = CCur(Trim$(listView.ListItems(i).SubItems(4)))
        ufConfirmation.txtTotalFacturesS�lectionn�es.value = _
            Format$(ufConfirmation.txtTotalFacturesS�lectionn�es.value + valeur, "###,##0.00 $")
        ufConfirmation.txtNbFacturesS�lectionn�es.value = _
            ufConfirmation.txtNbFacturesS�lectionn�es.value + 1
    Next i
    
    If ufConfirmation.txtNbFacturesS�lectionn�es.value = 1 Then
        ufConfirmation.cmdConfirmation.Caption = "Confirmer cette facture"
    Else
        ufConfirmation.cmdConfirmation.Caption = "Confirmer les (" & _
         ufConfirmation.txtNbFacturesS�lectionn�es.value & ") factures s�lectionn�es"
    End If
    ufConfirmation.cmdConfirmation.Visible = True
    
End Sub

Sub DecocherToutesLesCases(listView As listView) '2025-03-12 @ 12:40

    Dim i As Integer
    For i = 1 To listView.ListItems.count
        listView.ListItems(i).Checked = False
        Call MarquerLigneSelectionnee(listView.ListItems(i))
    Next i
    
    ufConfirmation.txtTotalFacturesS�lectionn�es = Format$(0, "###,##0.00 $")
    ufConfirmation.txtNbFacturesS�lectionn�es = 0
    ufConfirmation.cmdConfirmation.Visible = False
    
End Sub

Public Sub MarquerLigneSelectionnee(item As listItem) '2025-03-12 @ 12:40

    'V�rifie si l'�l�ment n'a pas d�j� la mention "   - S�lectionn�e -"
    If InStr(item.SubItems(3), "   - S�lectionn�e -") = 0 Then
        item.SubItems(3) = Left$(item.SubItems(3), 60) & "   - S�lectionn�e -"
    Else
        item.SubItems(3) = Left$(item.SubItems(3), 60)
    End If
    
End Sub

Sub Confirmation_Mise_�_Jour() '2025-03-12 @ 12:40

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
                ufConfirmation.txtNoFactureEnConfirmation.value = invNo
                DoEvents
                Call MAJ_Statut_Facture_Ent�te_BD_MASTER(invNo)
                Call MAJ_Statut_Facture_Ent�te_Local(invNo)
                DoEvents
                Call Construire_GL_Posting_Confirmation(invNo)
                DoEvents
            End If
        Next i
    End With

    MsgBox "La confirmation des factures est compl�t�e", vbOKOnly + vbInformation, "Confirmation de traitement"

    Unload ufConfirmation
    Call Afficher_ufConfirmation
    
End Sub

Sub MAJ_Statut_Facture_Ent�te_BD_MASTER(invoice As String) '2025-03-12 @ 12:40

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Ent�te_BD_MASTER", invoice, 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Ent�te$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
              ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim SQL As String
    'Open the recordset for the specified invoice
    SQL = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & invoice & "'"
    rs.Open SQL, conn, 2, 3
    If Not rs.EOF Then
        'Update AC_ouC with 'C'
        rs.Fields(fFacEACouC - 1).value = "C"
        rs.Update
    Else
        'Handle the case where the specified invoice is not found
        MsgBox "La facture '" & invoice & "' n'existe pas!", vbCritical
    End If
    
    'Close recordset and connection
    rs.Close
    conn.Close
    
    Application.ScreenUpdating = True

    'Lib�rer la m�moire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Ent�te_BD_MASTER", "", startTime)

End Sub

Sub MAJ_Statut_Facture_Ent�te_Local(invoice As String) '2025-03-12 @ 12:40
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Ent�te_Local", invoice, 0)
    
    Dim ws As Worksheet: Set ws = wsdFAC_Ent�te
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=invoice, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim r As Long, rowToBeUpdated As Long, tecID As Long
    If Not foundRange Is Nothing Then
        r = foundRange.row
        ws.Cells(r, fFacEACouC).value = "C"
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Ent�te."
    End If
    
    'Lib�rer la m�moire
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Ent�te_Local", "", startTime)

End Sub

Sub Construire_GL_Posting_Confirmation(invoice As String) '2025-03-12 @ 12:42

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:Construire_GL_Posting_Confirmation", invoice, 0)

    Dim ws As Worksheet: Set ws = wsdFAC_Ent�te
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=invoice, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim r As Long
    If Not foundRange Is Nothing Then
        r = foundRange.row
        Dim dateFact As Date
        dateFact = Left$(ws.Cells(r, fFacEDateFacture).value, 10)
        Dim hono As Currency
        hono = ws.Cells(r, fFacEHonoraires).value
        Dim misc1 As Currency, misc2 As Currency, misc3 As Currency
        misc1 = ws.Cells(r, fFacEAutresFrais1).value
        misc2 = ws.Cells(r, fFacEAutresFrais2).value
        misc3 = ws.Cells(r, fFacEAutresFrais3).value
        Dim tps As Currency, tvq As Currency
        tps = ws.Cells(r, fFacEMntTPS).value
        tvq = ws.Cells(r, fFacEMntTVQ).value
        
        Dim descGL_Trans As String, source As String
        descGL_Trans = ws.Cells(r, fFacENomClient).value
        source = "FACTURE:" & invoice
        
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
            MyArray(6, 1) = ObtenirNoGlIndicateur("TPS Factur�e")
            MyArray(6, 2) = "TPS percues"
            MyArray(6, 3) = -tps
            MyArray(6, 4) = ""
        End If
        
        'PST to pay (tvq)
        If tvq Then
            MyArray(7, 1) = ObtenirNoGlIndicateur("TVQ Factur�e")
            MyArray(7, 2) = "TVQ percues"
            MyArray(7, 3) = -tvq
            MyArray(7, 4) = ""
        End If
        
        'Mise � jour du posting GL des confirmations de facture
        Dim GLEntryNo As Long
        Call GL_Posting_To_DB(dateFact, descGL_Trans, source, MyArray, GLEntryNo)
        Call GL_Posting_Locally(dateFact, descGL_Trans, source, MyArray, GLEntryNo)
        
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Ent�te.", vbCritical
    End If
    
    'Lib�rer la m�moire
    On Error Resume Next
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Call Log_Record("modFAC_Confirmation:Construire_GL_Posting_Confirmation", "", startTime)

End Sub

