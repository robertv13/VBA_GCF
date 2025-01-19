Attribute VB_Name = "modFAC_Confirmation"
Option Explicit

Public invNo As String
Public Factures As Collection

Sub Afficher_ufConfirmation() '2025-01-19 @ 08:42

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

Sub PrepareDonneesPourListView() '2025-01-19 @ 08:42

    Set Factures = New Collection
    
    Call ObtenirFactureAConfirmer("AC")
    
    Dim ws As Worksheet
    Set ws = wshFAC_Entête
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AY").End(xlUp).row
    
    Dim invNo As String, dateFacture As String, nomClient As String, totalFacture As String
    Dim r As Long
    If lastUsedRow > 2 Then
        For r = 3 To lastUsedRow
            invNo = " " & ws.Range("AY" & r).Value
            dateFacture = " " & Format$(ws.Range("AZ" & r), wshAdmin.Range("B1").Value)
            nomClient = ws.Range("BC" & r).Value
            totalFacture = Format$(ws.Range("BO" & r).Value, "###,##0.00 $")
            totalFacture = Space(13 - Len(totalFacture)) & totalFacture
            Factures.Add Array(invNo, dateFacture, nomClient, totalFacture)
        Next r
    End If

End Sub

Sub ObtenirFactureAConfirmer(AC_OR_C As String) '2025-01-19 @ 08:42

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirFactureAConfirmer", 0)
    
    'Utilisation de la feuille FAC_Entête
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'Utilisation du AF#2 dans wshFAC_Entête
    
    'Effacer les données de la dernière utilisation
    ws.Range("AW6:AW10").ClearContents
    ws.Range("AW6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_FAC_Entête[#All]")
    ws.Range("AW7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("AW2:AW3")
    ws.Range("AW3").Value = AC_OR_C
    ws.Range("AW8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("AY1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("AY2:BP2")
    ws.Range("AW9").Value = rngResult.Address
        
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Qu'avons-nous comme résultat ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AY").End(xlUp).row
    ws.Range("AW10").Value = lastUsedRow - 2 & " lignes"
    
    If lastUsedRow > 3 Then
        With ws.Sort 'Sort - Inv_No
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("AY3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Invoice Number
            .SetRange ws.Range("AY3:BP" & lastUsedRow) 'Set Range
            .Apply 'Apply Sort
         End With
     End If

    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing

    Call Log_Record("modFAC_Confirmation:ObtenirFactureAConfirmer", startTime)

End Sub

Sub CocherToutesLesCases(listView As listView) '2025-01-19 @ 08:42

    'On s'assure de commencer avec aucune ligne de sélectionnée
    ufConfirmation.txtNbFacturesSélectionnées.Value = 0
    ufConfirmation.txtTotalFacturesSélectionnées.Value = 0
    
    Dim valeur As Currency
    Dim i As Integer
    For i = 1 To listView.ListItems.count
        listView.ListItems(i).Checked = True
        Call MarquerLigneSelectionnee(listView.ListItems(i))
        valeur = CCur(Trim(listView.ListItems(i).SubItems(4)))
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

Sub DecocherToutesLesCases(listView As listView) '2025-01-19 @ 08:42

    Dim i As Integer
    For i = 1 To listView.ListItems.count
        listView.ListItems(i).Checked = False
        Call MarquerLigneSelectionnee(listView.ListItems(i))
    Next i
    
    ufConfirmation.txtTotalFacturesSélectionnées = Format$(0, "###,##0.00 $")
    ufConfirmation.txtNbFacturesSélectionnées = 0
    ufConfirmation.cmdConfirmation.Visible = False
    
End Sub

Public Sub MarquerLigneSelectionnee(item As listItem)

    'Vérifie si l'élément n'a pas déjà la mention "   - Sélectionnée -"
    If InStr(item.SubItems(3), "   - Sélectionnée -") = 0 Then
        item.SubItems(3) = Left(item.SubItems(3), 60) & "   - Sélectionnée -"
    Else
        item.SubItems(3) = Left(item.SubItems(3), 60)
    End If
    
End Sub

Sub Confirmation_Mise_À_Jour() '2025-01-19 @ 08:42

    Dim Ligne As listItem
    
    ufConfirmation.lblFactureEmConfirmation.Visible = True
    ufConfirmation.txtNoFactureEnConfirmation.Visible = True

    Application.ScreenUpdating = True
    
    With ufConfirmation.ListView1
        Dim i As Long
        'Parcourir chacune des lignes
        For i = 1 To .ListItems.count
            Set Ligne = .ListItems(i)
            If Ligne.Checked Then
                invNo = Trim(Ligne.SubItems(1))
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

Sub MAJ_Statut_Facture_Entête_BD_MASTER(invoice As String) '2025-01-19 @ 08:42

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Entête_BD_MASTER(" & invoice & ")", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Entête$"
    
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
    
    Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Entête_BD_MASTER", startTime)

End Sub

Sub MAJ_Statut_Facture_Entête_Local(invoice As String) '2025-01-19 @ 08:42
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Entête_Local(" & invoice & ")", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
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
    
    Call Log_Record("modFAC_Confirmation:MAJ_Statut_Facture_Entête_Local", startTime)

End Sub

Sub Construire_GL_Posting_Confirmation(invoice As String) '2024-08-18 @17:15

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:Construire_GL_Posting_Confirmation(" & invoice & ")", 0)

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
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
        dateFact = Left(ws.Cells(r, fFacEDateFacture).Value, 10)
        Dim hono As Currency
        hono = ws.Cells(r, fFacEHonoraires).Value
        Dim misc1 As Currency, misc2 As Currency, misc3 As Currency
        misc1 = ws.Cells(r, fFacEAutresFrais1).Value
        misc2 = ws.Cells(r, fFacEAutresFrais2).Value
        misc3 = ws.Cells(r, fFacEAutresFrais3).Value
        Dim tps As Currency, tvq As Currency
        tps = ws.Cells(r, fFacEMntTPS).Value
        tvq = ws.Cells(r, fFacEMntTVQ).Value
        
        Dim descGL_Trans As String, source As String
        descGL_Trans = ws.Cells(r, fFacENomClient).Value
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
        Call GL_Posting_To_DB(dateFact, descGL_Trans, source, MyArray, GLEntryNo)
        Call GL_Posting_Locally(dateFact, descGL_Trans, source, MyArray, GLEntryNo)
        
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Entête.", vbCritical
    End If
    
    'Libérer la mémoire
    On Error Resume Next
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Call Log_Record("modFAC_Confirmation:Construire_GL_Posting_Confirmation", startTime)

End Sub


Sub ObtenirFactureInfos(noFact As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirFactureInfos", 0)
    
    'Save original worksheet
    Dim oWorkSheet As Worksheet: Set oWorkSheet = ActiveSheet
    
    'Reference to A/R master file
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim result As Variant
    Dim rngToSearch As Range: Set rngToSearch = ws.Range("A1").CurrentRegion.offset(0, 0).Resize(lastUsedRow, 1)
    result = Application.WorksheetFunction.XLookup(noFact, _
                                                   rngToSearch, _
                                                   rngToSearch, _
                                                   "Not Found", _
                                                   0, _
                                                   1)
    
    If result <> "Not Found" Then
        Dim matchedRow As Long
        matchedRow = Application.Match(noFact, rngToSearch, 0)
        
        Call AfficherInformationsFacture(ws, matchedRow)
        
        Call AfficherPDFetWIPicones
        
        Dim resultArr As Variant
        resultArr = Fn_ObtenirTECFacturésPourFacture(noFact)
        
        If Not IsEmpty(resultArr) Then
            Dim TECSummary() As Variant
            ReDim TECSummary(1 To 10, 1 To 3)
            Call ObtenirSommaireTEC(resultArr, TECSummary)
            
            Dim FeesSummary() As Variant
            ReDim FeesSummary(1 To 5, 1 To 3)
            Call ObtenirSommaireDesTaux(resultArr, FeesSummary)
        End If
        oWorkSheet.Activate
    Else
        MsgBox "La facture n'existe pas"
        GoTo CleanExit
    End If
    
CleanExit:
    Set oWorkSheet = Nothing
    Set rngToSearch = Nothing
    Set ws = Nothing

    Call Log_Record("modFAC_Confirmation:ObtenirFactureInfos", startTime)

End Sub

Sub AfficherPDFetWIPicones()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:AfficherPDFetWIPicones", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    Dim i As Long
    Dim iconPath As String
    iconPath = wshAdmin.Range("F5").Value & Application.PathSeparator & "Resources"
    
    Dim pic As Picture
    Dim cell As Range
    
    '1. Insert the PDF icon
    
    'Set the cell where the icon should be inserted
    Set cell = ws.Cells(7, 12) 'Set the cell where the icon should be inserted
            
    Set pic = ws.Pictures.Insert(iconPath & Application.PathSeparator & "AdobeAcrobatReader.png")
    With pic
        .Name = "PDF"
        .Top = cell.Top + 10
        .Left = cell.Left + 10
        .Height = 50 'cell.Height
        .Width = 50 'cell.width
        .Placement = xlMoveAndSize
        .OnAction = "shpPDF_Click"
    End With
    
    '2. Insert the WIP icon
    
    'Set the cell where the icon should be inserted
    Set cell = ws.Cells(14, 5) 'Set the cell where the icon should be inserted
    
    Set pic = ws.Pictures.Insert(iconPath & Application.PathSeparator & "WIP.png")
    With pic
        .Name = "WIP"
        .Top = cell.Top + 10
        .Left = cell.Left + 10
        .Height = 50 'cell.Height
        .Width = 50 'cell.width
        .Placement = xlMoveAndSize
        .OnAction = "shpWIP_Click"
    End With
    
    'Libérer la mémoire
    Set cell = Nothing
    Set pic = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Confirmation:AfficherPDFetWIPicones", startTime)
    
End Sub

Sub AfficherInformationsFacture(wsF As Worksheet, r As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:AfficherInformationsFacture", 0)
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    'Display all fields from FAC_Entête
    With ws
        .Range("L5").Value = wsF.Cells(r, 2).Value
    
        ws.Range("F7").Value = wsF.Cells(r, 5).Value
        ws.Range("F8").Value = wsF.Cells(r, 6).Value
        ws.Range("F9").Value = wsF.Cells(r, 7).Value
        ws.Range("F10").Value = wsF.Cells(r, 8).Value
        ws.Range("F11").Value = wsF.Cells(r, 9).Value
        
        ws.Range("L13").Value = wsF.Cells(r, 10).Value
        ws.Range("L14").Value = wsF.Cells(r, 12).Value
        ws.Range("L15").Value = wsF.Cells(r, 14).Value
        ws.Range("L16").Value = wsF.Cells(r, 16).Value
        ws.Range("L17").formula = "=SUM(L13:L16)"
        
        ws.Range("L18").Value = wsF.Cells(r, 18).Value
        ws.Range("L19").Value = wsF.Cells(r, 20).Value
        ws.Range("L21").formula = "=SUM(L17:L19)"
        
        ws.Range("L23").Value = wsF.Cells(r, 22).Value
        ws.Range("L25").formula = "=L21 - L23"
        
    End With
    
    'Take care of invoice type (to be confirmed OR already confirmed)
    If wsF.Cells(r, 3).Value = "AC" Then
        ws.Range("H5").Value = "À CONFIRMER"
        ws.Shapes("shpConfirmerFacture").Visible = True
    Else
        ws.Range("H5").Value = ""
        ws.Shapes("shpConfirmerFacture").Visible = False
    End If
    
    'Make OK button visible
    ws.Shapes("shpOK").Visible = True
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Application.EnableEvents = True

    Call Log_Record("modFAC_Confirmation:AfficherFactureFormatPDF", startTime)

End Sub

Sub shpWIP_Click()

    Call ObtenirListeTECFacturés
    
End Sub

Sub ObtenirListeTECFacturés()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirListeTECFacturés", 0)
    
    'Utilisation d'un AdvancedFilter directement dans TEC_Local (BI:BX)
    Call ObtenirListeTECFacturésFiltreAvancé(invNo)

    Dim ws As Worksheet: Set ws = wshTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "BJ").End(xlUp).row
    
    'Est-ce que nous avons des TEC pour cette facture ?
    If lastUsedRow < 3 Then
        MsgBox "Il n'y a aucun TEC associé à la facture '" & invNo & "'"
    Else
        Call PreparerRapportTECFacturés
    End If
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modFAC_Confirmation:ObtenirListeTECFacturés", startTime)
    
End Sub

Sub PreparerRapportTECFacturés()

'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:PreparerRapportTECFacturés", 0)
    
    'Assigner la feuille du rapport
    Dim strRapport As String
    strRapport = "Rapport TEC facturés"
    Dim wsRapport As Worksheet: Set wsRapport = wshTECFacturé
    wsRapport.Cells.Clear
    
    'Désactiver les mises à jour de l'écran et autres alertes
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    'Mettre en forme la feuille de rapport
    With wsRapport
        ' Titre du rapport
        .Range("A1").Value = "TEC facturés pour la facture '" & invNo & "'"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.size = 12
        
        'Ajouter une date de génération du rapport
        .Range("A2").Value = "Date de création : " & Format(Now, "dd/mm/yyyy")
        .Range("A2").Font.Italic = True
        .Range("A2").Font.size = 10
        
        'Entête du rapport (A4:D4)
        .Range("A4").Value = "Date"
        .Range("B4").Value = "Prof."
        .Range("C4").Value = "Description"
        .Range("D4").Value = "Heures"
        With .Range("A4:D4")
            .Font.size = 9
            .Font.Bold = True
            .Font.Italic = True
            .Font.Color = vbWhite
            .HorizontalAlignment = xlCenter
        End With
        
        'Utilisation du AdvancedFilter # 3 sur la feuille TEC_Local
        Dim wsSource As Worksheet
        Set wsSource = wshTEC_Local 'Utilisation des résultats du AF (BJ:BY)
        
        'Copier quelques données de la source
        Dim rngResult As Range
        Set rngResult = wsSource.Range("BJ1").CurrentRegion.offset(2, 0)
        'Redimensionner la plage après l'offset pour avoir que les données (pas d'entête)
        Set rngResult = rngResult.Resize(rngResult.Rows.count - 2)
        'Transfert des données vers un tableau
        Dim tableau As Variant
        tableau = rngResult.Value
        
        'Créer un tableau pour les résultats
        Dim output() As Variant
        ReDim output(1 To UBound(tableau, 1), 1 To 4)
        Dim r As Long
        
        Dim i As Long
        For i = LBound(tableau, 1) To UBound(tableau, 1)
            r = r + 1
            output(r, 1) = tableau(i, 4)
            output(r, 2) = tableau(i, 3)
            output(r, 3) = tableau(i, 7)
            output(r, 4) = tableau(i, 8)
        Next i

        'Copier le tableau dans la feuille du rapport  partir de la ligne 5, colonne 1
        .Range(.Cells(5, 1), .Cells(5 + UBound(output, 1) - 1, 1 + UBound(output, 2) - 1)).Value = output
        'Ligne dans la feuille du rapport
        r = 5 + UBound(output, 1) - 1
        
        'Corps du rapport
        .Range("A5:D" & r).VerticalAlignment = xlCenter
        With .Range("A4:D4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        'Ajouter une bordure aux données
        .Range("A4:D" & r).Borders.LineStyle = xlContinuous
        With .Range("A5:D" & r).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Range("A5:D" & r).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        
        .Range("A4:D" & r).Font.Name = "Aptos Narrow"
        .Range("A4:D" & r).Font.size = 10
        
        .Columns("A").ColumnWidth = 10
        .Range("A4:A" & r).HorizontalAlignment = xlCenter
        
        .Columns("B").ColumnWidth = 6
        .Range("B4:B" & r).HorizontalAlignment = xlCenter
        
        .Columns("C").ColumnWidth = 72
        .Columns("C").WrapText = True
        
        .Columns("D").ColumnWidth = 7
        .Columns("D").NumberFormat = "##0.00"
        
    End With

    'Configurer la mise en page pour l'impression ou l'export en PDF
    With wsRapport.PageSetup
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        
        'Ajuster la marge des en-têtes et pieds de page (1 cm)
        .HeaderMargin = Application.CentimetersToPoints(1)
        .FooterMargin = Application.CentimetersToPoints(1)
        
        .Orientation = xlPortrait 'Portrait
        .FitToPagesWide = 1 'Ajuster sur une page en largeur
        .FitToPagesTall = False ' Ne pas ajuster en hauteur
        .PrintArea = "A1:D" & r ' Définir la zone d'impression
        .CenterHorizontally = True ' Centrer horizontalement
        .CenterVertically = False ' Centrer verticalement
    End With
    
    'On se déplace à la feuille contenant le rapport
    wsRapport.Visible = xlSheetVisible
    wsRapport.Activate
    
    MsgBox "Le rapport a été généré sur la feuille " & strRapport
    
    'Libérer la mémoire
    Set rngResult = Nothing
    Set wsRapport = Nothing
    Set wsSource = Nothing
    
'    Call Log_Record("modFAC_Confirmation:PreparerRapportTECFacturés", startTime)
    
End Sub

Sub ObtenirListeTECFacturésFiltreAvancé(noFact As String) '2024-10-20 @ 11:11

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirListeTECFacturésFiltreAvancé", 0)

    'Utilisation de la feuille TEC_Local
    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    'wshTEC_Local_AF#3
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'AdvancedFilter par Numéro de Facture
    
    'Effacer les données de la dernière utilisation
    ws.Range("BH6:BH10").ClearContents
    ws.Range("BH6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_TEC_Local[#All]")
    ws.Range("BH7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("BH2:BH3")
    ws.Range("BH3").Value = CStr(noFact)
    ws.Range("BH8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("BJ1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("BJ2:BY2")
    ws.Range("BH9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Qu'avons-nous comme résultat ?
    Dim lastResultRow As Long
    lastResultRow = ws.Cells(ws.Rows.count, "BJ").End(xlUp).row
    ws.Range("BH10").Value = lastResultRow - 2 & " lignes"
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 3 Then
        With ws.Sort 'Sort - Date, ProfID, TECID
            .SortFields.Clear
            'First sort On Date
            .SortFields.Add key:=ws.Range("BM3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On ProfID
            .SortFields.Add key:=ws.Range("BK3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.Add key:=ws.Range("BJ3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange ws.Range("BJ3:BY" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'Free memory
    Set rngData = Nothing
    Set rngCriteria = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Confirmation:ObtenirListeTECFacturésFiltreAvancé", startTime)
    
End Sub

Sub ObtenirSommaireTEC(arr As Variant, ByRef TECSummary As Variant)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirSommaireTEC", 0)
    
    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    'Setup a Dictionary to summarize the hours by Professionnal
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")

    Dim pro As String
    Dim hres As Double
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        pro = wsTEC.Cells(arr(i), 3).Value
        hres = wsTEC.Cells(arr(i), 8).Value
        If hres <> 0 Then
            If dictHours.Exists(pro) Then
                dictHours(pro) = dictHours(pro) + hres
            Else
                dictHours.Add pro, hres
            End If
        End If
    Next i
    
    Dim profID As Long
    Dim rowInWorksheet As Long: rowInWorksheet = 13
    Dim prof As Variant
    Application.EnableEvents = False
    If dictHours.count <> 0 Then
        For Each prof In Fn_Sort_Dictionary_By_Value(dictHours, True) 'Sort dictionary by hours in descending order
            Dim strProf As String
            strProf = prof
            profID = Fn_GetID_From_Initials(strProf)
            hres = dictHours(prof)
            Dim tauxHoraire As Currency
            tauxHoraire = Fn_Get_Hourly_Rate(profID, wshFAC_Confirmation.Range("L5").Value)
            wshFAC_Confirmation.Cells(rowInWorksheet, 6) = strProf
            wshFAC_Confirmation.Cells(rowInWorksheet, 7) = _
                    CDbl(Format$(hres, "0.00"))
            wshFAC_Confirmation.Cells(rowInWorksheet, 8) = _
                    CDbl(Format$(tauxHoraire, "# ##0.00 $"))
            rowInWorksheet = rowInWorksheet + 1
    '        Debug.Print "#054 - Summary : " & strProf & " = " & hres & " @ " & tauxHoraire
    '        Cells(rowSelected, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
    '        rowSelected = rowSelected + 1
        Next prof
    End If
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set dictHours = Nothing
    Set prof = Nothing
    Set wsTEC = Nothing
    
    Call Log_Record("modFAC_Confirmation:ObtenirSommaireTEC", startTime)

End Sub

Sub ObtenirTotalTEC(arr As Variant, ByRef TECTotal As Double)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirTotalTEC", 0)
    
    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    'Setup a Dictionary to summarize the hours by Professionnal
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")

    Dim pro As String
    Dim hres As Double
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        pro = wsTEC.Cells(arr(i), 3).Value
        hres = wsTEC.Cells(arr(i), 8).Value
        If hres <> 0 Then
            If dictHours.Exists(pro) Then
                dictHours(pro) = dictHours(pro) + hres
            Else
                dictHours.Add pro, hres
            End If
        End If
    Next i
    
    Dim profID As Long
    Dim rowInWorksheet As Long: rowInWorksheet = 13
    Dim prof As Variant
    Application.EnableEvents = False
    If dictHours.count <> 0 Then
        For Each prof In dictHours
            Dim strProf As String
            strProf = prof
            profID = Fn_GetID_From_Initials(strProf)
            hres = dictHours(prof)
            Dim tauxHoraire As Currency
            tauxHoraire = Fn_Get_Hourly_Rate(profID, wshFAC_Confirmation.Range("L5").Value)
            wshFAC_Confirmation.Cells(rowInWorksheet, 6) = strProf
            wshFAC_Confirmation.Cells(rowInWorksheet, 7) = _
                    CDbl(Format$(hres, "0.00"))
            wshFAC_Confirmation.Cells(rowInWorksheet, 8) = _
                    CDbl(Format$(tauxHoraire, "# ##0.00 $"))
            rowInWorksheet = rowInWorksheet + 1
    '        Debug.Print "#055 - Summary : " & strProf & " = " & hres & " @ " & tauxHoraire
    '        Cells(rowSelected, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
    '        rowSelected = rowSelected + 1
        Next prof
    End If
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set dictHours = Nothing
    Set prof = Nothing
    Set wsTEC = Nothing
    
    Call Log_Record("modFAC_Confirmation:ObtenirTotalTEC", startTime)

End Sub

Sub ObtenirSommaireDesTaux(arr As Variant, ByRef FeesSummary As Variant)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:ObtenirSommaireDesTaux", 0)
    
    Dim wsFees As Worksheet: Set wsFees = wshFAC_Sommaire_Taux
    
    'Determine the last used row
    Dim lastUsedRow As Long
    lastUsedRow = wsFees.Cells(wsFees.Rows.count, 1).End(xlUp).row
    
    'Get Invoice number
    Dim invNo As String
    invNo = Trim(wshFAC_Confirmation.Range("L5").Value)
    
    'Use Range.Find to locate the first cell with the InvoiceNo
    Dim cell As Range
    Set cell = wsFees.Range("A2:A" & lastUsedRow).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the invNo was found at all
    Dim firstAddress As String
    Dim rowFeesSummary As Long: rowFeesSummary = 20
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Application.EnableEvents = False
        Do
            'Display values in the worksheet
            If wsFees.Cells(cell.row, 4).Value <> 0 Then
                wshFAC_Confirmation.Range("F" & rowFeesSummary).Value = wsFees.Cells(cell.row, 3).Value
                wshFAC_Confirmation.Range("G" & rowFeesSummary).Value = _
                            CCur(Format$(wsFees.Cells(cell.row, 4).Value, "##0.00"))
                wshFAC_Confirmation.Range("H" & rowFeesSummary).Value = _
                            CCur(Format$(wsFees.Cells(cell.row, 5).Value, "##,##0.00 $"))
                rowFeesSummary = rowFeesSummary + 1
            End If
            'Find the next cell with the invNo
            Set cell = wsFees.Range("A2:A" & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        Application.EnableEvents = True
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set wsFees = Nothing
    
    Call Log_Record("modFAC_Confirmation:ObtenirSommaireDesTaux", startTime)

End Sub

Sub NettoyerCellulesEtIconesPDF()

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshFAC_Confirmation:NettoyerCellulesEtIconesPDF", 0)
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    Application.ScreenUpdating = False
    
    ws.Range("F5:J5,L5,F7:I11,L13:L19,L21,L23,L25,F13:H17,F20:H24").ClearContents
    
    Dim pic As Picture
    For Each pic In ws.Pictures
        On Error Resume Next
        pic.Delete
        On Error GoTo 0
    Next pic
    
    Application.ScreenUpdating = True
    
    'Hide both buttons
    ws.Shapes("shpConfirmerFacture").Visible = False
    ws.Shapes("shpOK").Visible = False
    
    'Libérer la mémoire
    Set pic = Nothing
    Set ws = Nothing

    Application.EnableEvents = True
    
    wshFAC_Confirmation.Range("L5").Select
    
    Call Log_Record("modFAC_Confirmation:NettoyerCellulesEtIconesPDF", startTime)

End Sub

Sub ObtenirPostingExistantGL(invNo)

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshFAC_Confirmation:ObtenirPostingExistantGL", 0)
    
    Dim wsGL As Worksheet: Set wsGL = wshGL_Trans
    
    Dim lastUsedRow
    lastUsedRow = wsGL.Cells(wsGL.Rows.count, "A").End(xlUp).row
    Dim rngToSearch As Range: Set rngToSearch = wsGL.Range("D1:D" & lastUsedRow)
    
    'Use Range.Find to locate the first cell with the invNo
    Dim cell As Range
    Set cell = wsGL.Range("D2:D" & lastUsedRow).Find(What:="FACTURE:" & invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the invNo was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Dim r As Long
        r = 38
        Application.EnableEvents = False
        Do
            'Save the information for invoice deletion
            r = r + 1
            'Find the next cell with the invNo
            Set cell = wsGL.Range("D2:D" & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        Application.EnableEvents = True
    End If

    'Libérer la mémoire
    Set cell = Nothing
    Set rngToSearch = Nothing
    Set wsGL = Nothing
    
    Call Log_Record("modFAC_Confirmation:ObtenirPostingExistantGL", startTime)

End Sub

Sub shpExit_Click()

    Call RetournerMenuFAC
    
End Sub

Sub RetournerMenuFAC()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:RetournerMenuFAC", 0)
   
    wshFAC_Confirmation.Unprotect '2024-08-21 @ 05:06
    
    Application.EnableEvents = False
    wshFAC_Confirmation.Range("F5:J5").ClearContents
    wshFAC_Confirmation.Range("L5").ClearContents
    Application.EnableEvents = True
    
    wshFAC_Confirmation.Visible = xlSheetHidden

    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select
    
    Call Log_Record("modFAC_Confirmation:RetournerMenuFAC", startTime)
    
End Sub

Sub shpExitDetailTEC_Click()

    ActiveSheet.Visible = xlSheetHidden
    wshFAC_Confirmation.Activate
    Call NettoyerCellulesEtIconesPDF

End Sub
