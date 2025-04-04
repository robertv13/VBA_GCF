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

'CommentOut - 2025-03-12 @ 12:43
'Sub AfficherPDFetWIPicones()
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:AfficherPDFetWIPicones", "", 0)
'
'    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
'
'    Dim i As Long
'    Dim iconPath As String
'    iconPath = wsdADMIN.Range("F5").value & Application.PathSeparator & "Resources"
'
'    Dim pic As Picture
'    Dim cell As Range
'
'    '1. Insert the PDF icon
'
'    'Set the cell where the icon should be inserted
'    Set cell = ws.Cells(7, 12) 'Set the cell where the icon should be inserted
'
'    Set pic = ws.Pictures.Insert(iconPath & Application.PathSeparator & "AdobeAcrobatReader.png")
'    With pic
'        .Name = "PDF"
'        .Top = cell.Top + 10
'        .Left = cell.Left + 10
'        .Height = 50 'cell.Height
'        .Width = 50 'cell.width
'        .Placement = xlMoveAndSize
'        .OnAction = "shpPDF_Click"
'    End With
'
'    '2. Insert the WIP icon
'
'    'Set the cell where the icon should be inserted
'    Set cell = ws.Cells(14, 5) 'Set the cell where the icon should be inserted
'
'    Set pic = ws.Pictures.Insert(iconPath & Application.PathSeparator & "WIP.png")
'    With pic
'        .Name = "WIP"
'        .Top = cell.Top + 10
'        .Left = cell.Left + 10
'        .Height = 50 'cell.Height
'        .Width = 50 'cell.width
'        .Placement = xlMoveAndSize
'        .OnAction = "shpWIP_Click"
'    End With
'
'    'Lib�rer la m�moire
'    Set cell = Nothing
'    Set pic = Nothing
'    Set ws = Nothing
'
'    Call Log_Record("modFAC_Confirmation:AfficherPDFetWIPicones", "", startTime)
'
'End Sub
'

'CommentOut - 2025-03-12 @ 12:44
'Sub AfficherInformationsFacture(wsF As Worksheet, r As Long)
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:AfficherInformationsFacture", "", 0)
'
'    Application.EnableEvents = False
'
'    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
'
'    'Display all fields from FAC_Ent�te
'    With ws
'        .Range("L5").value = wsF.Cells(r, 2).value
'
'        ws.Range("F7").value = wsF.Cells(r, 5).value
'        ws.Range("F8").value = wsF.Cells(r, 6).value
'        ws.Range("F9").value = wsF.Cells(r, 7).value
'        ws.Range("F10").value = wsF.Cells(r, 8).value
'        ws.Range("F11").value = wsF.Cells(r, 9).value
'
'        ws.Range("L13").value = wsF.Cells(r, 10).value
'        ws.Range("L14").value = wsF.Cells(r, 12).value
'        ws.Range("L15").value = wsF.Cells(r, 14).value
'        ws.Range("L16").value = wsF.Cells(r, 16).value
'        ws.Range("L17").formula = "=SUM(L13:L16)"
'
'        ws.Range("L18").value = wsF.Cells(r, 18).value
'        ws.Range("L19").value = wsF.Cells(r, 20).value
'        ws.Range("L21").formula = "=SUM(L17:L19)"
'
'        ws.Range("L23").value = wsF.Cells(r, 22).value
'        ws.Range("L25").formula = "=L21 - L23"
'
'    End With
'
'    'Take care of invoice type (to be confirmed OR already confirmed)
'    If wsF.Cells(r, 3).value = "AC" Then
'        ws.Range("H5").value = "� CONFIRMER"
'        ws.Shapes("shpConfirmerFacture").Visible = True
'    Else
'        ws.Range("H5").value = ""
'        ws.Shapes("shpConfirmerFacture").Visible = False
'    End If
'
'    'Make OK button visible
'    ws.Shapes("shpOK").Visible = True
'
'    'Lib�rer la m�moire
'    Set ws = Nothing
'
'    Application.EnableEvents = True
'
'    Call Log_Record("modFAC_Confirmation:AfficherInformationsFacture", "", startTime)
'
'End Sub
'

'CommentOut - 2025-03-12 @ 12:44
'Sub shpWIP_Click()
'
'    Call ObtenirListeTECFactur�s
'
'End Sub
'

'CommentOut - 2025-03-12 @ 12:44
'Sub ObtenirListeTECFactur�s()
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirListeTECFactur�s", "", 0)
'
'    'Utilisation d'un AdvancedFilter directement dans TEC_Local (BI:BX)
'    Call ObtenirListeTECFactur�sFiltreAvanc�(invNo)
'
'    Dim ws As Worksheet: Set ws = wsdTEC_Local
'    Dim lastUsedRow As Long
'    lastUsedRow = ws.Cells(ws.Rows.count, "BJ").End(xlUp).row
'
'    'Est-ce que nous avons des TEC pour cette facture ?
'    If lastUsedRow < 3 Then
'        MsgBox "Il n'y a aucun TEC associ� � la facture '" & invNo & "'"
'    Else
'        Call PreparerRapportTECFactures
'    End If
'
'    'Lib�rer la m�moire
'    Set ws = Nothing
'
'    Call Log_Record("modFAC_Confirmation:ObtenirListeTECFactur�s", "", startTime)
'
'End Sub
'

'CommentOut - 2025-03-12 @ 12:49
'Sub ObtenirListeTECFactur�sFiltreAvanc�(noFact As String) '2024-10-20 @ 11:11
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirListeTECFactur�sFiltreAvanc�", "", 0)
'
'    'Utilisation de la feuille TEC_Local
'    Dim ws As Worksheet: Set ws = wsdTEC_Local
'
'    'wsdTEC_Local_AF#3
'
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'
'    'AdvancedFilter par Num�ro de Facture
'
'    'Effacer les donn�es de la derni�re utilisation
'    ws.Range("BH6:BH10").ClearContents
'    ws.Range("BH6").value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'D�finir le range pour la source des donn�es en utilisant un tableau
'    Dim rngData As Range
'    Set rngData = ws.Range("l_tbl_TEC_Local[#All]")
'    ws.Range("BH7").value = rngData.Address
'
'    'D�finir le range des crit�res
'    Dim rngCriteria As Range
'    Set rngCriteria = ws.Range("BH2:BH3")
'    ws.Range("BH3").value = CStr(noFact)
'    ws.Range("BH8").value = rngCriteria.Address
'
'    'D�finir le range des r�sultats et effacer avant le traitement
'    Dim rngResult As Range
'    Set rngResult = ws.Range("BJ1").CurrentRegion
'    rngResult.offset(2, 0).Clear
'    Set rngResult = ws.Range("BJ2:BY2")
'    ws.Range("BH9").value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Qu'avons-nous comme r�sultat ?
'    Dim lastResultRow As Long
'    lastResultRow = ws.Cells(ws.Rows.count, "BJ").End(xlUp).row
'    ws.Range("BH10").value = lastResultRow - 2 & " lignes"
'
'    'Est-il n�cessaire de trier les r�sultats ?
'    If lastResultRow > 3 Then
'        With ws.Sort 'Sort - Date, ProfID, TECID
'            .SortFields.Clear
'            'First sort On Date
'            .SortFields.Add key:=ws.Range("BM3"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Second, sort On ProfID
'            .SortFields.Add key:=ws.Range("BK3"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Third, sort On TecID
'            .SortFields.Add key:=ws.Range("BJ3"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            .SetRange ws.Range("BJ3:BY" & lastResultRow)
'            .Apply 'Apply Sort
'         End With
'    End If
'
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'
'    'Free memory
'    Set rngData = Nothing
'    Set rngCriteria = Nothing
'    Set rngResult = Nothing
'    Set ws = Nothing
'
'    Call Log_Record("modFAC_Confirmation:ObtenirListeTECFactur�sFiltreAvanc�", "", startTime)
'
'End Sub
'

'CommentOut - 2025-03-12 @ 12:50
'Sub ObtenirSommaireTEC(arr As Variant, ByRef TECSummary As Variant)
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirSommaireTEC", "", 0)
'
'    Dim wsTEC As Worksheet: Set wsTEC = wsdTEC_Local
'
'    'Setup a Dictionary to summarize the hours by Professionnal
'    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
'
'    Dim pro As String
'    Dim hres As Double
'    Dim i As Long
'    For i = 1 To UBound(arr, 1)
'        pro = wsTEC.Cells(arr(i), 3).value
'        hres = wsTEC.Cells(arr(i), 8).value
'        If hres <> 0 Then
'            If dictHours.Exists(pro) Then
'                dictHours(pro) = dictHours(pro) + hres
'            Else
'                dictHours.Add pro, hres
'            End If
'        End If
'    Next i
'
'    Dim profID As Long
'    Dim rowInWorksheet As Long: rowInWorksheet = 13
'    Dim prof As Variant
'    Application.EnableEvents = False
'    If dictHours.count <> 0 Then
'        For Each prof In Fn_Sort_Dictionary_By_Value(dictHours, True) 'Sort dictionary by hours in descending order
'            Dim strProf As String
'            strProf = prof
'            profID = Fn_GetID_From_Initials(strProf)
'            hres = dictHours(prof)
'            Dim tauxHoraire As Currency
'            tauxHoraire = Fn_Get_Hourly_Rate(profID, wshFAC_Confirmation.Range("L5").value)
'            wshFAC_Confirmation.Cells(rowInWorksheet, 6) = strProf
'            wshFAC_Confirmation.Cells(rowInWorksheet, 7) = _
'                    CDbl(Format$(hres, "0.00"))
'            wshFAC_Confirmation.Cells(rowInWorksheet, 8) = _
'                    CDbl(Format$(tauxHoraire, "# ##0.00 $"))
'            rowInWorksheet = rowInWorksheet + 1
'    '        Debug.Print "#054 - Summary : " & strProf & " = " & hres & " @ " & tauxHoraire
'    '        Cells(rowSelected, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
'    '        rowSelected = rowSelected + 1
'        Next prof
'    End If
'    Application.EnableEvents = True
'
'    'Lib�rer la m�moire
'    Set dictHours = Nothing
'    Set prof = Nothing
'    Set wsTEC = Nothing
'
'    Call Log_Record("modFAC_Confirmation:ObtenirSommaireTEC", "", startTime)
'
'End Sub
'

'CommenOut - 2025-03-12 @ 12:50
'Sub ObtenirTotalTEC(arr As Variant, ByRef TECTotal As Double)
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirTotalTEC", "", 0)
'
'    Dim wsTEC As Worksheet: Set wsTEC = wsdTEC_Local
'
'    'Setup a Dictionary to summarize the hours by Professionnal
'    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
'
'    Dim pro As String
'    Dim hres As Double
'    Dim i As Long
'    For i = 1 To UBound(arr, 1)
'        pro = wsTEC.Cells(arr(i), 3).value
'        hres = wsTEC.Cells(arr(i), 8).value
'        If hres <> 0 Then
'            If dictHours.Exists(pro) Then
'                dictHours(pro) = dictHours(pro) + hres
'            Else
'                dictHours.Add pro, hres
'            End If
'        End If
'    Next i
'
'    Dim profID As Long
'    Dim rowInWorksheet As Long: rowInWorksheet = 13
'    Dim prof As Variant
'    Application.EnableEvents = False
'    If dictHours.count <> 0 Then
'        For Each prof In dictHours
'            Dim strProf As String
'            strProf = prof
'            profID = Fn_GetID_From_Initials(strProf)
'            hres = dictHours(prof)
'            Dim tauxHoraire As Currency
'            tauxHoraire = Fn_Get_Hourly_Rate(profID, wshFAC_Confirmation.Range("L5").value)
'            wshFAC_Confirmation.Cells(rowInWorksheet, 6) = strProf
'            wshFAC_Confirmation.Cells(rowInWorksheet, 7) = _
'                    CDbl(Format$(hres, "0.00"))
'            wshFAC_Confirmation.Cells(rowInWorksheet, 8) = _
'                    CDbl(Format$(tauxHoraire, "# ##0.00 $"))
'            rowInWorksheet = rowInWorksheet + 1
'    '        Debug.Print "#055 - Summary : " & strProf & " = " & hres & " @ " & tauxHoraire
'    '        Cells(rowSelected, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
'    '        rowSelected = rowSelected + 1
'        Next prof
'    End If
'    Application.EnableEvents = True
'
'    'Lib�rer la m�moire
'    Set dictHours = Nothing
'    Set prof = Nothing
'    Set wsTEC = Nothing
'
'    Call Log_Record("modFAC_Confirmation:ObtenirTotalTEC", "", startTime)
'
'End Sub
'

'CommentOut - 2025-03-12 @ 12:51
'Sub ObtenirSommaireDesTaux(arr As Variant, ByRef FeesSummary As Variant)
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirSommaireDesTaux", "", 0)
'
'    Dim wsFees As Worksheet: Set wsFees = wsdFAC_Sommaire_Taux
'
'    'Determine the last used row
'    Dim lastUsedRow As Long
'    lastUsedRow = wsFees.Cells(wsFees.Rows.count, 1).End(xlUp).row
'
'    'Get Invoice number
'    Dim invNo As String
'    invNo = Trim$(wshFAC_Confirmation.Range("L5").value)
'
'    'Use Range.Find to locate the first cell with the InvoiceNo
'    Dim cell As Range
'    Set cell = wsFees.Range("A2:A" & lastUsedRow).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
'
'    'Check if the invNo was found at all
'    Dim firstAddress As String
'    Dim rowFeesSummary As Long: rowFeesSummary = 20
'    If Not cell Is Nothing Then
'        firstAddress = cell.Address
'        Application.EnableEvents = False
'        Do
'            'Display values in the worksheet
'            If wsFees.Cells(cell.row, 4).value <> 0 Then
'                wshFAC_Confirmation.Range("F" & rowFeesSummary).value = wsFees.Cells(cell.row, 3).value
'                wshFAC_Confirmation.Range("G" & rowFeesSummary).value = _
'                            CCur(Format$(wsFees.Cells(cell.row, 4).value, "##0.00"))
'                wshFAC_Confirmation.Range("H" & rowFeesSummary).value = _
'                            CCur(Format$(wsFees.Cells(cell.row, 5).value, "##,##0.00 $"))
'                rowFeesSummary = rowFeesSummary + 1
'            End If
'            'Find the next cell with the invNo
'            Set cell = wsFees.Range("A2:A" & lastUsedRow).FindNext(After:=cell)
'        Loop While Not cell Is Nothing And cell.Address <> firstAddress
'        Application.EnableEvents = True
'    End If
'
'    'Lib�rer la m�moire
'    Set cell = Nothing
'    Set wsFees = Nothing
'
'    Call Log_Record("modFAC_Confirmation:ObtenirSommaireDesTaux", "", startTime)
'
'End Sub
'
'Sub NettoyerCellulesEtIconesPDF()
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:NettoyerCellulesEtIconesPDF", "", 0)
'
'    Application.EnableEvents = False
'
'    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
'
'    Application.ScreenUpdating = False
'
'    ws.Range("F5:J5,L5,F7:I11,L13:L19,L21,L23,L25,F13:H17,F20:H24").ClearContents
'
'    Dim pic As Picture
'    For Each pic In ws.Pictures
'        On Error Resume Next
'        pic.Delete
'        On Error GoTo 0
'    Next pic
'
'    Application.ScreenUpdating = True
'
'    'Hide both buttons
'    ws.Shapes("shpConfirmerFacture").Visible = False
'    ws.Shapes("shpOK").Visible = False
'
'    'Lib�rer la m�moire
'    Set pic = Nothing
'    Set ws = Nothing
'
'    Application.EnableEvents = True
'
'    wshFAC_Confirmation.Range("L5").Select
'
'    Call Log_Record("modFAC_Confirmation:NettoyerCellulesEtIconesPDF", "", startTime)
'
'End Sub
'

'CommentOut - 2025-03-12 @ 12:53
'Sub ObtenirPostingExistantGL(invNo As String)
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirPostingExistantGL", "", 0)
'
'    Dim wsGL As Worksheet: Set wsGL = wsdGL_Trans
'
'    Dim lastUsedRow As Long
'    lastUsedRow = wsGL.Cells(wsGL.Rows.count, "A").End(xlUp).row
'    Dim rngToSearch As Range: Set rngToSearch = wsGL.Range("D1:D" & lastUsedRow)
'
'    'Use Range.Find to locate the first cell with the invNo
'    Dim cell As Range
'    Set cell = wsGL.Range("D2:D" & lastUsedRow).Find(What:="FACTURE:" & invNo, LookIn:=xlValues, LookAt:=xlWhole)
'
'    'Check if the invNo was found at all
'    Dim firstAddress As String
'    If Not cell Is Nothing Then
'        firstAddress = cell.Address
'        Dim r As Long
'        r = 38
'        Application.EnableEvents = False
'        Do
'            'Save the information for invoice deletion
'            r = r + 1
'            'Find the next cell with the invNo
'            Set cell = wsGL.Range("D2:D" & lastUsedRow).FindNext(After:=cell)
'        Loop While Not cell Is Nothing And cell.Address <> firstAddress
'        Application.EnableEvents = True
'    End If
'
'    'Lib�rer la m�moire
'    Set cell = Nothing
'    Set rngToSearch = Nothing
'    Set wsGL = Nothing
'
'    Call Log_Record("modFAC_Confirmation:ObtenirPostingExistantGL", "", startTime)
'
'End Sub
'

'CommentOut - 2025-03-12 @ 12:54
'Sub shpExit_Click()
'
'    Call RetournerMenuFAC
'
'End Sub
'

'CommentOut - 2025-03-12 @ 12:54
'Sub RetournerMenuFAC()
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:RetournerMenuFAC", "", 0)
'
'    wshFAC_Confirmation.Unprotect '2024-08-21 @ 05:06
'
'    Application.EnableEvents = False
'    wshFAC_Confirmation.Range("F5:J5").ClearContents
'    wshFAC_Confirmation.Range("L5").ClearContents
'    Application.EnableEvents = True
'
'    wshFAC_Confirmation.Visible = xlSheetHidden
'
'    wshMenuFAC.Activate
'    wshMenuFAC.Range("A1").Select
'
'    Call Log_Record("modFAC_Confirmation:RetournerMenuFAC", "", startTime)
'
'End Sub
'
