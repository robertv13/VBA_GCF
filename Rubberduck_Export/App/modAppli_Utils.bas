Attribute VB_Name = "modAppli_Utils"
'@Folder("Général")

Option Explicit

'@Description "Structure pour VerifierTEC"
Public Type StatistiquesTEC '2025-06-19 @ 10:41
    cas_doublon_TecID As Long
    cas_date_invalide As Long
    cas_date_future As Long
    cas_hres_invalide As Long
    cas_estFacturable_invalide As Long
    cas_estFacturee_invalide As Long
    cas_date_fact_invalide As Long
    cas_date_facture_future As Long
    cas_estDetruit_invalide As Long
    nbValid As Long
    nbInvalid As Long
    totalHeures As Double
    total_hres_inscrites As Double
    total_hres_detruites As Double
    total_hres_facturees As Double
    total_hres_facturables As Double
    total_hres_non_facturables As Double
End Type

Public Sub ConvertirPlageABooleen(rng As Range)

    Dim cell As Range
    
    For Each cell In rng
        Select Case cell.Value
            Case 0, "False", "FAUX" 'False
                cell.Value = "FAUX"
            Case -1, "True", "VRAI" 'True
                cell.Value = "VRAI"
            Case Else
                MsgBox cell.Value & " est une valeur INVALIDE pour la cellule '" & cell.Address & "'" & vbNewLine & vbNewLine & _
                    "Veuillez contacter le développeur sans faute", _
                    "Erreur de logique", _
                    vbCritical
        End Select
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Sub

Sub MettreEnFormeImpressionSimple(ws As Worksheet, rng As Range, header1 As String, _
                       header2 As String, titleRows As String, Optional Orient As String = "L")
    
    On Error GoTo CleanUp
    
    Application.PrintCommunication = False
    
    With ws.PageSetup
        .PrintArea = rng.Address
        .PrintTitleRows = titleRows
        .PrintTitleColumns = vbNullString
        
        .CenterHeader = "&""-,Gras""&12&K0070C0" & header1 & Chr$(10) & "&11" & header2
        
        .LeftFooter = "&8&D - &T"
        .CenterFooter = "&8&KFF0000&A"
        .RightFooter = "&""Segoe UI,Normal""&8Page &P of &N"
        
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1.5)
        
        .BottomMargin = Application.CentimetersToPoints(1)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        
        .CenterHorizontally = True
        
        If Orient = "L" Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .PaperSize = xlPaperLetter
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
CleanUp:
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
    
End Sub

Public Sub TransfererTableau2DVersPlage(ByRef arr As Variant, _
                               ByVal rngTo As Range, _
                               Optional ByVal clearExistingData As Boolean = True, _
                               Optional ByVal HeaderSize As Long = 1)
                        
    'Si requis, on efface le contenu de rngTo avant
    If clearExistingData = True Then
        rngTo.CurrentRegion.offset(HeaderSize).ClearContents
    End If
    
    'En fonction des dimensions du tableau (arr)
    Dim r As Long
    Dim c As Long

    r = UBound(arr, 1) - LBound(arr, 1) + HeaderSize
    c = UBound(arr, 2) - LBound(arr, 2) + HeaderSize
    rngTo.Resize(r, c).Value = arr
    
End Sub

Sub TransfererPlageVersTableau2D(ByVal rng As Range, ByRef arr As Variant, Optional ByVal headerRows As Long = 1)

    'La plage est-elle valide ?
    If rng Is Nothing Then
        MsgBox "La plage est invalide ou non définie.", vbExclamation, , "modAppli_Utils:TransfererPlageVersTableau2D"
        Exit Sub
    End If
    
    'Calculer la taille de la plage des données pour ensuite ignorer les en-têtes
    Dim numRows As Long
    Dim numCols As Long
    
    numRows = rng.Rows.count - headerRows
    numCols = rng.Columns.count
    
    'La plage contient-elle des données ?
    If numRows <= 0 Or numCols <= 0 Then
        MsgBox "Aucune donnée à copier dans le tableau.", vbExclamation, "modAppli_Utils:TransfererPlageVersTableau2D"
        Exit Sub
    End If
    
    'Définir la taille de la plage qui contient les données, en fonction de numRows & numCols
    On Error Resume Next
    Dim rngData As Range
    Set rngData = rng.Resize(numRows, numCols).Offset(headerRows, 0)
    On Error GoTo 0
    
    'Copier les données du Rage vers le tableau (Array)
    If Not rngData Is Nothing Then
        arr = rngData.Value
    Else
        MsgBox "Erreur lors de la création de la plage de données.", vbExclamation, "modAppli_Utils:TransfererPlageVersTableau2D"
    End If
    
    'Libérer la mémoire
    Set rngData = Nothing
    
End Sub

Sub CreerOuRemplacerFeuille(wsName As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:CreerOuRemplacerFeuille", vbNullString, 0)
    
    Dim wsExists As Boolean
    wsExists = Fn_FeuilleExiste(wsName)
    
    'Si la feuille existe, on la supprime
    If wsExists Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(wsName).Delete
        Application.DisplayAlerts = True
    End If

    'Attendre un instant pour éviter les conflits éventuels
    DoEvents
    
    'Ajouter une nouvelle feuille et la renommer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=wshMenu)
    ws.Name = wsName

    'Libérer la mémoire
    Set ws = Nothing

    Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:CreerOuRemplacerFeuille", vbNullString, startTime)
    
End Sub

Sub CreerEnteteDeFeuille(r As Range, couleurFond As Long) '2025-06-30 @ 14:08

    On Error GoTo GestionErreur

    'La plage 'r' est-elle valide ?
    If r Is Nothing Then Exit Sub

    'Vérifier que la couleur est numérique (couleur RGB valide)
    If Not IsNumeric(couleurFond) Or couleurFond < 0 Then Exit Sub

    With r
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = couleurFond
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .size = 9
            .Italic = True
            .Bold = True
        End With
        .HorizontalAlignment = xlCenter
    End With

    'Auto-ajustement des colonnes concernées
    Dim ws As Worksheet
    Set ws = r.Worksheet
    ws.Columns(r.Columns(1).Column).Resize(, r.Columns.count).AutoFit

    Exit Sub

GestionErreur:
    MsgBox "Erreur dans CreerEnteteDeFeuille : " & Err.description, vbExclamation, "Erreur VBA"
    
End Sub

Sub AjouterMessageAuxResultats(ws As Worksheet, r As Long, c As Long, m As String)

    ws.Cells(r, c).Value = m
    If c = 1 Then
        ws.Cells(r, c).Font.Bold = True
    End If

End Sub

Sub AppliquerConditionalFormating(rng As Range, headerRows As Long, couleurFond As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:AppliquerConditionalFormating", vbNullString, 0)
    
    'Avons-nous un Range valide ?
    If rng Is Nothing Or rng.Rows.count <= headerRows Then
        Exit Sub
    End If
    
    Dim dataRange As Range
    
    'Définir la plage de données à laquelle appliquer la mise en forme conditionnelle, en
    'excluant les lignes d'en-tête
    Set dataRange = rng.Resize(rng.Rows.count - headerRows).Offset(headerRows, 0)
    
    'Effacer les formats conditionnels existants sur la plage de données
    dataRange.Interior.ColorIndex = xlNone

    'Appliquer les couleurs en alternance
    Dim i As Long
    For i = 1 To dataRange.Rows.count
        'Vérifier la position réelle de la ligne dans la feuille
        If (dataRange.Rows(i).Row + headerRows) Mod 2 = 0 Then
            dataRange.Rows(i).Interior.Color = couleurFond
        End If
    Next i
    
    'Libérer la mémoire
    Set dataRange = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:AppliquerConditionalFormating", vbNullString, startTime)

End Sub

Sub AppliquerFormatColonnesParTable(ws As Worksheet, rng As Range, HeaderRow As Long)

    Dim lo As ListObject
    Dim rngUnion As Range
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:AppliquerFormatColonnesParTable", vbNullString, 0)
    
    Select Case rng.Worksheet.CodeName
       Case "wsdCC_Regularisations" '2025-11-14 @ 16:31
            Set lo = ws.ListObjects("l_tbl_CC_Regularisations")
            Set rngUnion = Application.Union( _
                lo.ListColumns(fREGULClientNom).DataBodyRange, _
                lo.ListColumns(fREGULDescription).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementGauche(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fREGULRegulID).DataBodyRange, _
                lo.ListColumns(fREGULInvNo).DataBodyRange, _
                lo.ListColumns(fREGULDate).DataBodyRange, _
                lo.ListColumns(fREGULTimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fREGULHono).DataBodyRange, _
                lo.ListColumns(fREGULFrais).DataBodyRange, _
                lo.ListColumns(fREGULTPS).DataBodyRange, _
                lo.ListColumns(fREGULTVQ).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURR_DOLLARS)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fREGULInvNo).DataBodyRange, modFormats.FMT_DATE)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fREGULTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)

        Case "wsdDEB_Recurrent"  '2025-11-14 @ 09:53
            Set lo = ws.ListObjects("l_tbl_DEB_Recurrent")
            Call modFormats.AppliquerAlignementCentre(lo.ListColumns(fDebRNoDebRec).DataBodyRange)
            Call modFormats.AppliquerAlignementCentre(lo.ListColumns(fDebRTimeStamp).DataBodyRange)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fDebRDate).DataBodyRange, modFormats.FMT_DATE)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fDebRType).DataBodyRange, _
                lo.ListColumns(fDebRReference).DataBodyRange, _
                lo.ListColumns(fDebRCompte).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementGauche(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fDebRTotal).DataBodyRange, _
                lo.ListColumns(fDebRTPS).DataBodyRange, _
                lo.ListColumns(fDebRTVQ).DataBodyRange, _
                lo.ListColumns(fDebRCréditTPS).DataBodyRange, _
                lo.ListColumns(fDebRCréditTVQ).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURR_DOLLARS)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fDebRTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)

            'Hors tableau structuré (Sommaire)
            ws.Columns("P").NumberFormat = modFormats.FMT_ENTIER
            ws.Columns("R").HorizontalAlignment = xlRight
            ws.Columns("R").NumberFormat = modFormats.FMT_MNT_CURR_DOLLARS
            ws.Columns("S").NumberFormat = modFormats.FMT_DATE_HEURE
    
        Case "wsdDEB_Trans" '2025-11-14 @ 09:53
            Set lo = ws.ListObjects("l_tbl_DEB_Trans")
            Set rngUnion = Application.Union( _
                lo.ListColumns(fDebTType).DataBodyRange, _
                lo.ListColumns(fDebTBeneficiaire).DataBodyRange, _
                lo.ListColumns(fDebTDescription).DataBodyRange, _
                lo.ListColumns(fDebTReference).DataBodyRange, _
                lo.ListColumns(fDebTCompte).DataBodyRange, _
                lo.ListColumns(fDebTAutreRemarque).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementGauche(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fDebTNoEntrée).DataBodyRange, _
                lo.ListColumns(fDebTDate).DataBodyRange, _
                lo.ListColumns(fDebTFournID).DataBodyRange, _
                lo.ListColumns(fDebTNoCompte).DataBodyRange, _
                lo.ListColumns(fDebTCodeTaxe).DataBodyRange, _
                lo.ListColumns(fDebTTimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fDebTDate).DataBodyRange, modFormats.FMT_DATE)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fDebTTotal).DataBodyRange, _
                lo.ListColumns(fDebTTPS).DataBodyRange, _
                lo.ListColumns(fDebTTVQ).DataBodyRange, _
                lo.ListColumns(fDebTCréditTPS).DataBodyRange, _
                lo.ListColumns(fDebTCréditTVQ).DataBodyRange, _
                lo.ListColumns(fDebTDépense).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURR_DOLLARS)
            
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fDebTTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)

        Case "wsdENC_Details" '2025-11-14 @ 16:43
            Set lo = ws.ListObjects("l_tbl_ENC_Details")
            Call modFormats.AppliquerAlignementGauche(lo.ListColumns(fEncDCustomer).DataBodyRange)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fEncDPayID).DataBodyRange, _
                lo.ListColumns(fEncDInvNo).DataBodyRange, _
                lo.ListColumns(fEncDPayDate).DataBodyRange, _
                lo.ListColumns(fEncDTimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Call modFormats.AppliquerAlignementDroit(lo.ListColumns(fEncDPayAmount).DataBodyRange)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fEncDPayAmount).DataBodyRange, modFormats.FMT_MNT_CURRENCY)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fEncDPayDate).DataBodyRange, modFormats.FMT_DATE)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fEncDTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)

        Case "wsdENC_Entete" '2025-11-14 @ 16:50
            Set lo = ws.ListObjects("l_tbl_ENC_Entete")
            Set rngUnion = Application.Union( _
                lo.ListColumns(fEncECustomer).DataBodyRange, _
                lo.ListColumns(fEncEPayType).DataBodyRange, _
                lo.ListColumns(fEncENotes).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementGauche(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fEncEPayID).DataBodyRange, _
                lo.ListColumns(fEncEPayDate).DataBodyRange, _
                lo.ListColumns(fEncECodeClient).DataBodyRange, _
                lo.ListColumns(fEncETimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fEncEPayID).DataBodyRange, modFormats.FMT_ENTIER)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fEncEPayDate).DataBodyRange, modFormats.FMT_DATE)
            Call modFormats.AppliquerAlignementDroit(lo.ListColumns(fEncEAmount).DataBodyRange)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fEncEAmount).DataBodyRange, modFormats.FMT_MNT_CURR_DOLLARS)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fEncETimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)

        Case "wsdFAC_Comptes_Clients" '2025-11-14 @ 16:56
            Set lo = ws.ListObjects("l_tbl_FAC_Comptes_Clients")
            Call modFormats.AppliquerAlignementGauche(lo.ListColumns(fFacCCCustomer).DataBodyRange)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacCCInvNo).DataBodyRange, _
                lo.ListColumns(fFacCCInvoiceDate).DataBodyRange, _
                lo.ListColumns(fFacCCCodeClient).DataBodyRange, _
                lo.ListColumns(fFacCCStatus).DataBodyRange, _
                lo.ListColumns(fFacCCTerms).DataBodyRange, _
                lo.ListColumns(fFacCCDueDate).DataBodyRange, _
                lo.ListColumns(fFacCCDaysOverdue).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacCCTotal).DataBodyRange, _
                lo.ListColumns(fFacCCTotalPaid).DataBodyRange, _
                lo.ListColumns(fFacCCTotalRegul).DataBodyRange, _
                lo.ListColumns(fFacCCBalance).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURR_DOLLARS)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacCCInvoiceDate).DataBodyRange, _
                lo.ListColumns(fFacCCDueDate).DataBodyRange _
                )
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_DATE)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacCCDaysOverdue).DataBodyRange, modFormats.FMT_ENTIER)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacCCTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)

        Case "wsdFAC_Details" '2025-11-14 @ 10:13
            Set lo = ws.ListObjects("l_tbl_FAC_Details")
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacDInvNo).DataBodyRange, _
                lo.ListColumns(fFacDInvRow).DataBodyRange, _
                lo.ListColumns(fFacDTimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Call modFormats.AppliquerAlignementGauche(lo.ListColumns(fFacDDescription).DataBodyRange)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacDHeures).DataBodyRange, modFormats.FMT_MNT_CURRENCY)
            Call modFormats.AppliquerAlignementDroit(lo.ListColumns(fFacDHeures).DataBodyRange)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacDTaux).DataBodyRange, _
                lo.ListColumns(fFacDHonoraires).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURR_DOLLARS)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacDTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)
        
        Case "wsdFAC_Entete" '2025-11-14 @ 09:54
            Set lo = ws.ListObjects("l_tbl_FAC_Entete")
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacEContact).DataBodyRange, _
                lo.ListColumns(fFacENomClient).DataBodyRange, _
                lo.ListColumns(fFacEAdresse1).DataBodyRange, _
                lo.ListColumns(fFacEAdresse2).DataBodyRange, _
                lo.ListColumns(fFacEAdresse3).DataBodyRange, _
                lo.ListColumns(fFacEAF1Desc).DataBodyRange, _
                lo.ListColumns(fFacEAF2Desc).DataBodyRange, _
                lo.ListColumns(fFacEAF3Desc).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementGauche(rngUnion)
            
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacEInvNo).DataBodyRange, _
                lo.ListColumns(fFacEDateFacture).DataBodyRange, _
                lo.ListColumns(fFacEACouC).DataBodyRange, _
                lo.ListColumns(fFacECustID).DataBodyRange, _
                lo.ListColumns(fFacETauxTPS).DataBodyRange, _
                lo.ListColumns(fFacETauxTVQ).DataBodyRange, _
                lo.ListColumns(fFacETimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacEHonoraires).DataBodyRange, _
                lo.ListColumns(fFacEAutresFrais1).DataBodyRange, _
                lo.ListColumns(fFacEAutresFrais2).DataBodyRange, _
                lo.ListColumns(fFacEAutresFrais3).DataBodyRange, _
                lo.ListColumns(fFacEMntTPS).DataBodyRange, _
                lo.ListColumns(fFacEMntTVQ).DataBodyRange, _
                lo.ListColumns(fFacEARTotal).DataBodyRange, _
                lo.ListColumns(fFacEDépôt).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURR_DOLLARS)
            
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacEDateFacture).DataBodyRange, modFormats.FMT_DATE)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacETauxTPS).DataBodyRange, _
                lo.ListColumns(fFacETauxTVQ).DataBodyRange _
                )
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_TAUX_PCT_3)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacETimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)

        Case "wsdFAC_Projets_Details" '2025-11-14 @ 16:58
            Set lo = ws.ListObjects("l_tbl_FAC_Projets_Details")
            Call modFormats.AppliquerAlignementGauche(lo.ListColumns(fFacPDProjetID).DataBodyRange)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacPDProjetID).DataBodyRange, _
                lo.ListColumns(fFacPDClientID).DataBodyRange, _
                lo.ListColumns(fFacPDTECID).DataBodyRange, _
                lo.ListColumns(fFacPDProfID).DataBodyRange, _
                lo.ListColumns(fFacPDDate).DataBodyRange, _
                lo.ListColumns(fFacPDProf).DataBodyRange, _
                lo.ListColumns(fFacPDestDetruite).DataBodyRange, _
                lo.ListColumns(fFacPDTimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Call modFormats.AppliquerAlignementDroit(lo.ListColumns(fFacPDHeures).DataBodyRange)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacPDDate).DataBodyRange, modFormats.FMT_DATE)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacPDHeures).DataBodyRange, modFormats.FMT_MNT_CURRENCY)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacPDTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)

        Case "wsdFAC_Projets_Entete" '2025-11-14 @ 17:10
            Set lo = ws.ListObjects("l_tbl_FAC_Projets_Entete")
            Call modFormats.AppliquerAlignementGauche(lo.ListColumns(fFacPENomClient).DataBodyRange)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacPEProjetID).DataBodyRange, _
                lo.ListColumns(fFacPEClientID).DataBodyRange, _
                lo.ListColumns(fFacPEDate).DataBodyRange, _
                lo.ListColumns(fFacPEProf1).DataBodyRange, _
                lo.ListColumns(fFacPEProf2).DataBodyRange, _
                lo.ListColumns(fFacPEProf3).DataBodyRange, _
                lo.ListColumns(fFacPEProf4).DataBodyRange, _
                lo.ListColumns(fFacPEProf5).DataBodyRange, _
                lo.ListColumns(fFacPEestDetruite).DataBodyRange, _
                lo.ListColumns(fFacPETimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacPEHonoTotal).DataBodyRange, _
                lo.ListColumns(fFacPEHres1).DataBodyRange, _
                lo.ListColumns(fFacPETauxH1).DataBodyRange, _
                lo.ListColumns(fFacPEHono1).DataBodyRange, _
                lo.ListColumns(fFacPEHres2).DataBodyRange, _
                lo.ListColumns(fFacPETauxH2).DataBodyRange, _
                lo.ListColumns(fFacPEHono2).DataBodyRange, _
                lo.ListColumns(fFacPEHres3).DataBodyRange, _
                lo.ListColumns(fFacPETauxH4).DataBodyRange, _
                lo.ListColumns(fFacPEHono3).DataBodyRange, _
                lo.ListColumns(fFacPEHres4).DataBodyRange, _
                lo.ListColumns(fFacPETauxH4).DataBodyRange, _
                lo.ListColumns(fFacPEHono4).DataBodyRange, _
                lo.ListColumns(fFacPEHres5).DataBodyRange, _
                lo.ListColumns(fFacPETauxH5).DataBodyRange, _
                lo.ListColumns(fFacPEHono5).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURR_DOLLARS)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacPETauxH1).DataBodyRange, _
                lo.ListColumns(fFacPETauxH2).DataBodyRange, _
                lo.ListColumns(fFacPETauxH4).DataBodyRange, _
                lo.ListColumns(fFacPETauxH4).DataBodyRange, _
                lo.ListColumns(fFacPETauxH5).DataBodyRange _
                )
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURRENCY)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacPEDate).DataBodyRange, modFormats.FMT_DATE)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacPETimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)
            
        Case "wsdFAC_Sommaire_Taux" '2025-11-14 @ 17:18
            Set lo = ws.ListObjects("l_tbl_FAC_Sommaire_Taux")
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacSTInvNo).DataBodyRange, _
                lo.ListColumns(fFacSTSéquence).DataBodyRange, _
                lo.ListColumns(fFacSTProf).DataBodyRange, _
                lo.ListColumns(fFacSTTimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fFacSTHeures).DataBodyRange, _
                lo.ListColumns(fFacSTTaux).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacSTHeures).DataBodyRange, modFormats.FMT_MNT_CURRENCY)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacSTTaux).DataBodyRange, modFormats.FMT_MNT_CURR_DOLLARS)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fFacSTTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)
        
        Case "wsdGL_EJ_Recurrente" '2025-11-14 @ 17:23
            Set lo = ws.ListObjects("l_tbl_GL_EJ_Auto")
            Set rngUnion = Application.Union( _
                lo.ListColumns(fGlEjRDescription).DataBodyRange, _
                lo.ListColumns(fGlEjRCompte).DataBodyRange, _
                lo.ListColumns(fGlEjRAutreRemarque).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementGauche(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fGlEjRNoEjR).DataBodyRange, _
                lo.ListColumns(fGlEjRNoCompte).DataBodyRange, _
                lo.ListColumns(fGlEjRTimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fGlEjRDébit).DataBodyRange, _
                lo.ListColumns(fGlEjRCrédit).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURR_DOLLARS)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fGlEjRTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)

        Case "wsdGL_Trans" '2025-11-14 @ 17:27
            Set lo = ws.ListObjects("l_tbl_GL_Trans")
            Set rngUnion = Application.Union( _
                lo.ListColumns(fGlTDescription).DataBodyRange, _
                lo.ListColumns(fGlTSource).DataBodyRange, _
                lo.ListColumns(fGlTCompte).DataBodyRange, _
                lo.ListColumns(fGlTAutreRemarque).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementGauche(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fGlTNoEntrée).DataBodyRange, _
                lo.ListColumns(fGlTDate).DataBodyRange, _
                lo.ListColumns(fGlTNoCompte).DataBodyRange, _
                lo.ListColumns(fGlTTimeStamp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fGlTDébit).DataBodyRange, _
                lo.ListColumns(fGlTCrédit).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementDroit(rngUnion)
            Call modFormats.AppliquerNumberFormat(rngUnion, modFormats.FMT_MNT_CURR_DOLLARS)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fGlTNoEntrée).DataBodyRange, modFormats.FMT_ENTIER)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fGlTDate).DataBodyRange, modFormats.FMT_DATE)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fGlTTimeStamp).DataBodyRange, modFormats.FMT_DATE_HEURE)
        
        Case "wsdTEC_Local" '2025-11-14 @ 17:36
            Set lo = ws.ListObjects("l_tbl_TEC_Local")
            Set rngUnion = Application.Union( _
                lo.ListColumns(fTECClientNom).DataBodyRange, _
                lo.ListColumns(fTECDescription).DataBodyRange, _
                lo.ListColumns(fTECCommentaireNote).DataBodyRange, _
                lo.ListColumns(fTECVersionApp).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementGauche(rngUnion)
            Set rngUnion = Application.Union( _
                lo.ListColumns(fTECTECID).DataBodyRange, _
                lo.ListColumns(fTECProfID).DataBodyRange, _
                lo.ListColumns(fTECProf).DataBodyRange, _
                lo.ListColumns(fTECClientID).DataBodyRange, _
                lo.ListColumns(fTECEstFacturable).DataBodyRange, _
                lo.ListColumns(fTECDateSaisie).DataBodyRange, _
                lo.ListColumns(fTECEstFacturee).DataBodyRange, _
                lo.ListColumns(fTECDateFacturee).DataBodyRange, _
                lo.ListColumns(fTECEstDetruit).DataBodyRange, _
                lo.ListColumns(fTECNoFacture).DataBodyRange _
                )
            Call modFormats.AppliquerAlignementCentre(rngUnion)
            Call modFormats.AppliquerAlignementDroit(lo.ListColumns(fTECHeures).DataBodyRange)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fTECHeures).DataBodyRange, modFormats.FMT_MNT_CURRENCY)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fTECDate).DataBodyRange, modFormats.FMT_DATE)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fTECDateFacturee).DataBodyRange, modFormats.FMT_DATE)
            Call modFormats.AppliquerNumberFormat(lo.ListColumns(fTECDateSaisie).DataBodyRange, modFormats.FMT_DATE_HEURE)

    End Select

    'Post-traitements communs (AutoFit + RowHeight)
    If Not lo Is Nothing Then
        Call modFormats.AppliquerCommonPost(ws, lo)
    End If

    'Ajustements de largeurs de colonnes spécifiques
    Select Case rng.Worksheet.CodeName
        Case "wsdFAC_Entete" '2025-11-14 @ 17:51
            Call modFormats.AppliquerLargeurColonne(ws, lo.ListColumns(fFacENomClient).Range.Column, 50)
        Case "wsdTEC_Local"
            Call modFormats.AppliquerLargeurColonne(ws, lo.ListColumns(fTECClientNom).Range.Column, 45)
            Call modFormats.AppliquerLargeurColonne(ws, lo.ListColumns(fTECDescription).Range.Column, 60)
            Call modFormats.AppliquerLargeurColonne(ws, lo.ListColumns(fTECCommentaireNote).Range.Column, 30)
    End Select
    
    'Libérer la mémoire
    Set rngUnion = Nothing
    Set lo = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:AppliquerFormatColonnesParTable", vbNullString, startTime)

End Sub

Sub ObtenirDeplacementsAPartirDesTEC()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:ObtenirDeplacementsAPartirDesTEC", vbNullString, 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Mise en place de la feuille de sortie (output)
    Dim strOutput As String
    strOutput = "X_TEC_Déplacements"
    Call CreerOuRemplacerFeuille(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").Value = "Date"
    wsOutput.Range("B1").Value = "Date"
    wsOutput.Range("C1").Value = "Nom du client"
    wsOutput.Range("D1").Value = "Heures"
    wsOutput.Range("E1").Value = "Adresse_1"
    wsOutput.Range("F1").Value = "Adresse_2"
    wsOutput.Range("G1").Value = "Ville"
    wsOutput.Range("H1").Value = "Province"
    wsOutput.Range("I1").Value = "CodePostal"
    wsOutput.Range("J1").Value = "DistanceKM"
    wsOutput.Range("K1").Value = "Montant"
    Call CreerEnteteDeFeuille(wsOutput.Range("A1:K1"), RGB(0, 112, 192))
    
    'Analyse de TEC_Local
    Call modImport.ImporterClients
    Call modImport.ImporterTEC
    
    'Feuille pour les clients
    Dim wsMF As Worksheet: Set wsMF = wsdBD_Clients
    Dim lastUsedRowClientMF As Long
    lastUsedRowClientMF = wsMF.Cells(wsMF.Rows.count, 1).End(xlUp).Row
    Dim rngClientsMF As Range
    Set rngClientsMF = wsMF.Range("A1:A" & lastUsedRowClientMF)
    
    'Get From and To Dates
    Dim dateFrom As Date, dateTo As Date
    dateFrom = wsdADMIN.Range("MoisPrecDe").Value
    dateTo = wsdADMIN.Range("MoisPrecA").Value
    
    Dim wsTEC As Worksheet: Set wsTEC = wsdTEC_Local
    
    Dim lastUsedRowTEC As Long
    lastUsedRowTEC = wsTEC.Cells(wsTEC.Rows.count, 1).End(xlUp).Row
    Dim arr() As Variant
    
    'Copier le range en mémoire
    Call TransfererPlageVersTableau2D(wsTEC.Range("A1:P" & lastUsedRowTEC), arr, 2)
    
    'Mise en place d'un tableau pour recevoir les résultats (performance)
    Dim output() As Variant
    ReDim output(1 To UBound(arr, 1), 1 To UBound(arr, 2))
    Dim rowOutput As Long
    rowOutput = 1
    
    Dim clientData As Variant
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 3) = "GC" And UCase$(arr(i, 14)) <> "VRAI" Then
            If arr(i, 4) >= CLng(dateFrom) And arr(i, 4) <= CLng(dateTo) Then
                output(rowOutput, 1) = arr(i, 4)
                output(rowOutput, 2) = arr(i, 4)
                output(rowOutput, 4) = arr(i, 8)
                clientData = Fn_LigneClientAPartirDuClientID(Trim$(arr(i, 5)), wsMF)
                If IsArray(clientData) Then
                    output(rowOutput, 3) = clientData(1, fClntFMClientNom)
                    output(rowOutput, 5) = clientData(1, fClntFMAdresse1)
                    output(rowOutput, 6) = clientData(1, fClntFMAdresse2)
                    output(rowOutput, 7) = clientData(1, fClntFMVille)
                    output(rowOutput, 8) = clientData(1, fClntFMProvince)
                    output(rowOutput, 9) = clientData(1, fClntFMCodePostal)
                End If
                rowOutput = rowOutput + 1
            End If
        End If
    Next i
    
    'Copier le tableau dans le range
    Call TransfererTableau2DVersPlage(output, wsOutput.Range("A2:I" & UBound(output, 1)), True, 1)
    
    'Tri des données
    With wsOutput.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsOutput.Range("B2"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortTextAsNumbers 'Sort Date
        .SortFields.Add key:=wsdTEC_Local.Range("C2"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort on Client's name
        .SortFields.Add key:=wsdTEC_Local.Range("D2"), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal 'Sort on Hours
        .SetRange wsOutput.Range("A2:K" & rowOutput - 1) 'Set Range
        .Apply 'Apply Sort
     End With
    
    'Ajustement des formats
    With wsOutput
        .Range("A2:B" & rowOutput + 1).NumberFormat = wsdADMIN.Range("B1").Value
        .Range("D2:D" & rowOutput + 1).NumberFormat = "##0.00"
        .Range("A2:K" & rowOutput + 1).Font.Name = "Aptos Narrow"
        .Range("A2:K" & rowOutput + 1).Font.size = 10
        .Columns.AutoFit
    End With
    
    'Améliore le Look (saute 1 ligne entre chaque jour)
    For i = rowOutput To 3 Step -1
        If Len(Trim$(wsOutput.Cells(i, 3).Value)) > 0 Then
            If wsOutput.Cells(i, 2).Value <> wsOutput.Cells(i - 1, 2).Value Then
                wsOutput.Rows(i).Insert Shift:=xlDown
                wsOutput.Cells(i, 1).Value = wsOutput.Cells(i - 1, 2).Value
            End If
        End If
    Next i
    
    rowOutput = wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).Row
    
    'Améliore le Look (cache la date, le client et l'adresse si deux charges & +)
    Dim base As String
    For i = 2 To rowOutput
        If i = 2 Then
            base = wsOutput.Cells(i, 2).Value & wsOutput.Cells(i, 3).Value
        End If
        If i > 2 And Len(wsOutput.Cells(i, 2).Value) > 0 Then
            If wsOutput.Cells(i, 2).Value & wsOutput.Cells(i, 3).Value = base Then
                wsOutput.Cells(i, 2).Value = vbNullString
                wsOutput.Cells(i, 3).Value = vbNullString
                wsOutput.Cells(i, 5).Value = vbNullString
                wsOutput.Cells(i, 6).Value = vbNullString
                wsOutput.Cells(i, 7).Value = vbNullString
                wsOutput.Cells(i, 8).Value = vbNullString
                wsOutput.Cells(i, 9).Value = vbNullString
            Else
                base = wsOutput.Cells(i, 2).Value & wsOutput.Cells(i, 3).Value
            End If
        End If
    Next i
    
    'Result print setup - 2024-08-05 @ 05:16
    rowOutput = wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).Row
    
    For i = 3 To rowOutput
        If wsOutput.Cells(i, 1).Value > wsOutput.Cells(i - 1, 1).Value Then
            wsOutput.Cells(i, 2).Font.Bold = True
        Else
            wsOutput.Cells(i, 2).Value = vbNullString
        End If
    Next i
    
    'Première date est en caractère gras
    wsOutput.Cells(2, 2).Font.Bold = True
    rowOutput = rowOutput + 2
    wsOutput.Range("A" & rowOutput).Value = "**** " & Format$(lastUsedRowTEC - 2, "###,##0") & _
                                        " charges de temps analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("B2:K" & rowOutput)
    Call AppliquerConditionalFormating(rngArea, 1, RGB(173, 216, 230))

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Setup print parameters
    Dim header1 As String: header1 = "Liste des TEC pour Guillaume"
    Dim header2 As String: header2 = "Période du " & dateFrom & " au " & dateTo
    Call MettreEnFormeImpressionSimple(wsOutput, rngArea, header1, header2, "$1:$1", "P")
    
    'Libérer la mémoire
    Set rngArea = Nothing
    Set rngClientsMF = Nothing
    Set wsOutput = Nothing
    Set wsMF = Nothing
    Set wsTEC = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:ObtenirDeplacementsAPartirDesTEC", vbNullString, startTime)

End Sub

Sub ObtenirDateDernModifFichier(fileName As String, ByRef ddm As Date, _
                                    ByRef jours As Long, ByRef heures As Long, _
                                    ByRef minutes As Long, ByRef secondes As Long)
    
    'Créer une instance de FileSystemObject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Obtenir le fichier
    Dim fichier As Object: Set fichier = fso.GetFile(fileName)
    
    'Récupérer la date et l'heure de la dernière modification
    ddm = fichier.DateLastModified
    
    'Calculer la différence (jours) entre maintenant et la date de la dernière modification
    Dim diff As Double
    diff = Now - ddm
    
    'Convertir la différence en jours, heures, minutes et secondes
    jours = Int(diff)
    heures = Int((diff - jours) * 24)
    minutes = Int(((diff - jours) * 24 - heures) * 60)
    secondes = Int(((((diff - jours) * 24 - heures) * 60) - minutes) * 60)
    
    ' Libérer les objets
    Set fichier = Nothing
    Set fso = Nothing
    
End Sub

Sub RemplirPlageAvecCouleur(ByVal plage As Range, ByVal couleurRVB As Long)

    If Not plage Is Nothing Then
        Dim cellule As Range
        'Parcourt toutes les cellules de la plage (contiguës ou non)
        For Each cellule In plage
            On Error Resume Next
            Debug.Print "RemplirPlageAvecCouleur pour " & cellule.Address
            cellule.Interior.Color = couleurRVB
            On Error GoTo 0
        Next cellule
    Else
        MsgBox "La plage spécifiée est invalide.", vbExclamation, "Procédure 'RemplirPlageAvecCouleur'"
    End If
    
End Sub

Sub NoterNombreLignesParFeuille() '2025-01-22 @ 16:19

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:NoterNombreLignesParFeuille", vbNullString, 0)
    
    'Spécifiez les chemins des classeurs
    Dim cheminClasseurUsage As String
    cheminClasseurUsage = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & "GCF_File_Usage.xlsx"
    If Not Dir(cheminClasseurUsage, vbNormal) <> "" Then
        Call EnregistrerErreurs("modAppli_Utils", "NoterNombreLignesParFeuille", "Classeur d’usage introuvable", 0, "CRITICAL")
        Exit Sub
    End If

    Dim cheminClasseurMASTER As String
    cheminClasseurMASTER = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & wsdADMIN.Range("MASTER_FILE").Value
    
    Application.ScreenUpdating = False
    
    'Ouvrir le classeur d'usage
    Dim wbUsage As Workbook
    Set wbUsage = Workbooks.Open(cheminClasseurUsage)
    Dim wsUsage As Worksheet
    Set wsUsage = wbUsage.Worksheets("Data")
    
    'Trouver la première ligne disponible
    Dim LigneDisponible As Long
    LigneDisponible = wsUsage.Cells(wsUsage.Rows.count, 1).End(xlUp).Row + 1
    
    'Ouvrir le classeur maître en lecture seule
    Dim wbMaster As Workbook
    Set wbMaster = Workbooks.Open(cheminClasseurMASTER, ReadOnly:=True)
    
    'Ajouter l'horodatage à la première col
    Dim dateHeure As String
    dateHeure = Now
    wsUsage.Cells(LigneDisponible, 1).Value = Format$(dateHeure, "yyyy-mm-dd hh:nn:ss")
    
    'Parcourir les cols de la première ligne pour les noms de feuilles
    Dim feuilleNom As String
    Dim lastUsedRow As Long
    Dim col As Long
    col = 2 'Commence à la col 2
    Do While wsUsage.Cells(1, col).Value <> vbNullString
        feuilleNom = wsUsage.Cells(1, col).Value
        If Trim(feuilleNom) <> vbNullString Then
            'Vérifier si la feuille existe dans le classeur maître
            On Error Resume Next
            Dim wsMaster As Worksheet
            Set wsMaster = wbMaster.Sheets(feuilleNom)
            On Error GoTo 0
            
            lastUsedRow = 0
            If Not wsMaster Is Nothing Then
                lastUsedRow = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row
            End If
            
            'Écrire le résultat dans la ligne disponible
            wsUsage.Cells(LigneDisponible, col).Value = lastUsedRow
        End If
        'Passer à la col suivante
        col = col + 1
    Loop
    
    'Fermer le classeur maître sans enregistrer
    wbMaster.Close False
    
    Application.ScreenUpdating = True
    
    'Sauvegarder et fermer le classeur d'usage
    wbUsage.Close SaveChanges:=True
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli_Utils:NoterNombreLignesParFeuille", vbNullString, startTime)

End Sub

Function Fn_LireFichierTXT(chemin As String) As String '2025-08-12 @ 15:46

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fichier As Object
    
    If fso.fileExists(chemin) Then
        Set fichier = fso.OpenTextFile(chemin, 1)
        Fn_LireFichierTXT = fichier.ReadLine
        fichier.Close
    Else
        Fn_LireFichierTXT = ""
    End If
    
End Function

