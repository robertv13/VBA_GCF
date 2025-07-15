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

'Variables globales pour le module
Private gverificationIntegriteOK As Boolean
Private gsoldeComptesClients As Currency
Private gValeursAComparer() As Variant

'@Description("Routine pour ajouter des lignes de message à la feuille")
Private Sub AjouterMessage(ws As Worksheet, ByRef r As Long, ByRef c As Long, message As String)
Attribute AjouterMessage.VB_Description = "Routine pour ajouter des lignes de message à la feuille"

    ws.Cells(r, c).Value = message
    If c = 1 Then
        ws.Cells(r, c).Font.Bold = True
    End If
    r = r + 1
    
End Sub

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

Public Sub VerifierIntegriteTablesLocales() '2024-11-20 @ 06:55

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierIntegriteTablesLocales", vbNullString, 0)

    Application.ScreenUpdating = False
    
    'Variable pour déterminer à la fin s'il y a des erreurs...
    gverificationIntegriteOK = True
    
    Call modDev_Utils.EffacerEtRecreerWorksheet("X_Analyse_Integrite")
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Integrite")
    With wsOutput
        .Unprotect
        .Cells.Font.Name = "Courier New"
        .Cells.Font.size = 10
        .Range("A1").Value = "Feuille"
        .Range("B1").Value = "Message"
        .Range("C1").Value = "TimeStamp"
        .Columns("C").NumberFormat = wsdADMIN.Range("B1").Value & " hh:mm:ss"
        Call Make_It_As_Header(.Range("A1:C1"), RGB(0, 112, 192))
    End With

    'Data starts at row 2
    Dim r As Long: r = 2
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "Version du programme")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, ThisWorkbook.Name)
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1

    'Répertoire de données
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "Répertoire utilisé")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, wsdADMIN.Range("FolderSharedData").Value & gDATA_PATH)
    r = r + 1

    'Fichier utilisé
    Dim masterFileName As String
    masterFileName = "GCF_BD_MASTER.xlsx"
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "Fichier utilisé")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, masterFileName)
    r = r + 1
    
    'Date dernière modification du fichier MAÎTRE
    Dim fullFileName As String
    fullFileName = wsdADMIN.Range("FolderSharedData").Value & gDATA_PATH & Application.PathSeparator & masterFileName
    Dim ddm As Date
    Dim j As Long
    Dim h As Long
    Dim m As Long
    Dim s As Long

    Call Get_Date_Derniere_Modification(fullFileName, ddm, j, h, m, s)
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "Date dern. modification")
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = Format$(ddm, wsdADMIN.Range("B1").Value & " hh:mm:ss") & _
            " soit " & j & " jours, " & h & " heures, " & m & " minutes et " & s & " secondes"
    rng.Characters(1, 19).Font.Color = vbRed
    rng.Characters(1, 19).Font.Bold = True
    r = r + 2
    
    Dim readRows As Long
    'Tableau pour comparer les montants entre procédures
    ReDim gValeursAComparer(1 To 12, 1 To 4)
    
    'dnrPlanComptable ----------------------------------------------------- Plan Comptable
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "Plan Comptable")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Les informations ont été importées de la feuille 'Admin'")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierPlanComptable(wsOutput, r, readRows)

    'wsdBD_Clients --------------------------------------------------------------- Clients
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "BD_Clients")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "La feuille a été importée du fichier BD_Entrée.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierClients(wsOutput, r, readRows)
    
    'wsdBD_Fournisseurs ----------------------------------------------------- Fournisseurs
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "BD_Fournisseurs")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "La feuille a été importée du fichier BD_Entrée.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierFournisseurs(wsOutput, r, readRows)
    
    'wsdDEB_Recurrent ------------------------------------------------------ DEB_Récurrent
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "DEB_Récurrent")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "DEB_Récurrent a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierDEBRecurrent(wsOutput, r, readRows)
    
    'wsdDEB_Trans -------------------------------------------------------------- DEB_Trans
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "DEB_Trans")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "DEB_Trans a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierDEBTrans(wsOutput, r, readRows)
    
    'wsdFAC_Entete ------------------------------------------------------------ FAC_Entête
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "FAC_Entête")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "FAC_Entête a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierFACEntete(wsOutput, r, readRows)
    
    'wsdFAC_Details ---------------------------------------------------------- FAC_Détails
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "FAC_Détails")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "FAC_Détails a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierFACDetails(wsOutput, r, readRows)
    
    'wsdFAC_Comptes_Clients ------------------------------------------ FAC_Comptes_Clients
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "FAC_Comptes_Clients")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "FAC_Comptes_Clients a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierFACComptesClients(wsOutput, r, readRows)
    
    'wsdFAC_Sommaire_Taux ---------------------------------------------- FAC_Sommaire_Taux
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "FAC_Sommaire_Taux")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "FAC_Sommaire_Taux a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierFACSommaireTaux(wsOutput, r, readRows)
    
    'wsdENC_Entete ------------------------------------------------------------ ENC_Entête
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "ENC_Entête")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "ENC_Entête a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierENCEntete(wsOutput, r, readRows)
    
    'wsdENC_Details ---------------------------------------------------------- ENC_Détails
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "ENC_Détails")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "ENC_Détails a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierENCDetails(wsOutput, r, readRows)
    
    'wsdCC_Regularisations ---------------------------------------------------------- CC_Régularisations
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "CC_Régularisations")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "CC_Régularisations a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierCCRegularisations(wsOutput, r, readRows)
    
    'wsdFAC_Projets_Entete -------------------------------------------- FAC_Projets_Entête
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "FAC_Projets_Entête")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "FAC_Projets_Entête a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierFACProjetsEntete(wsOutput, r, readRows)
    
    'wsdFAC_Projets_Details ------------------------------------------ FAC_Projets_Détails
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "FAC_Projets_Détails")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "FAC_Projets_Détails a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierFACProjetsDetails(wsOutput, r, readRows)
    
    'wsdGL_Trans ---------------------------------------------------------------- GL_Trans
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "GL_Trans")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "GL_Trans a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierGLTrans(wsOutput, r, readRows)
    
    'wsdGL_EJ_Recurrente ------------------------------------------------ GL_EJ_Recurrente
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "GL_EJ_Recurrente")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "GL_EJ_Recurrente a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierGLEJRecurrente(wsOutput, r, readRows)
   
    'wshTEC_TdB_Data -------------------------------------------------------- TEC_TdB_Data
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "TEC_TdB_Data")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "TEC_TdB_Data a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierTECTdBData(wsOutput, r, readRows)
    
    'wsdTEC_Local -------------------------------------------------------------- TEC_Local
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "TEC_Local")
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "TEC_Local a été importée du fichier GCF_BD_MASTER.xlsx")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    r = r + 1
    
    Call VerifierTEC(wsOutput, r, readRows)
    
    'Dernière section de vérification
    Call AjouterMessageAuxResultats(wsOutput, r, 1, "Dernières vérifications")
    Call AjouterMessageAuxResultats(wsOutput, r, 3, Format$(Now(), "yyyy-mm-dd hh:mm:ss"))
    
    'Comparaison des valeurs - Total des factures (confirmées)
    If gValeursAComparer(1, 2) <> gValeursAComparer(1, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(1, 1)), _
                                                "FAC_Entete", CCur(gValeursAComparer(1, 2)), _
                                                "FAC_Comptes_Clients", CCur(gValeursAComparer(1, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If
    
    'Comparaison des valeurs - Total des factures (à confirmer)
    If gValeursAComparer(2, 2) <> gValeursAComparer(2, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(2, 1)), _
                                                "FAC_Entete", CCur(gValeursAComparer(2, 2)), _
                                                "FAC_Comptes_Clients", CCur(gValeursAComparer(2, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If
        
    'Comparaison des valeurs - Montant encaissé à date
    If gValeursAComparer(3, 2) <> gValeursAComparer(3, 3) Or _
        gValeursAComparer(3, 2) <> gValeursAComparer(3, 4) Or _
        gValeursAComparer(3, 3) <> gValeursAComparer(3, 4) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(3, 1)), _
                                                "FAC_Comptes_Clients", CCur(gValeursAComparer(3, 2)), _
                                                "ENC_Entete", CCur(gValeursAComparer(3, 3)), _
                                                "ENC_Details", CCur(gValeursAComparer(3, 4)))
        gverificationIntegriteOK = False
    End If
    
    'Comparaison des valeurs - Montant régularisé à date
    If gValeursAComparer(4, 2) <> gValeursAComparer(4, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(4, 1)), _
                                                "FAC_Comptes_Clients", CCur(gValeursAComparer(4, 2)), _
                                                "CC_Regularisations", CCur(gValeursAComparer(4, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If
    
    'Comparaison des valeurs - Solde à recevoir
    If gValeursAComparer(5, 2) <> gValeursAComparer(5, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(5, 1)), _
                                                "FAC_Comptes_Clients", CCur(gValeursAComparer(5, 2)), _
                                                "GL_Trans", CCur(gValeursAComparer(5, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If
    
    'Comparaison des valeurs - Heures saisies
    If gValeursAComparer(6, 2) <> gValeursAComparer(6, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(6, 1)), _
                                                "TEC_TdB_Data", CCur(gValeursAComparer(6, 2)), _
                                                "TEC_Local", CCur(gValeursAComparer(6, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If
    
    'Comparaison des valeurs - Heures détruites
    If gValeursAComparer(7, 2) <> gValeursAComparer(7, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(7, 1)), _
                                                "TEC_TdB_Data", CCur(gValeursAComparer(7, 2)), _
                                                "TEC_Local", CCur(gValeursAComparer(7, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If
    
    'Comparaison des valeurs - Heures NETTES
    If gValeursAComparer(8, 2) <> gValeursAComparer(8, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(8, 1)), _
                                                "TEC_TdB_Data", CCur(gValeursAComparer(8, 2)), _
                                                "TEC_Local", CCur(gValeursAComparer(8, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If

    'Comparaison des valeurs - Heures Non-Facturables
    If gValeursAComparer(9, 2) <> gValeursAComparer(9, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(9, 1)), _
                                                "TEC_TdB_Data", CCur(gValeursAComparer(9, 2)), _
                                                "TEC_Local", CCur(gValeursAComparer(9, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If

    'Comparaison des valeurs - Heures Facturables
    If gValeursAComparer(10, 2) <> gValeursAComparer(10, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(10, 1)), _
                                                "TEC_TdB_Data", CCur(gValeursAComparer(10, 2)), _
                                                "TEC_Local", CCur(gValeursAComparer(10, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If

    'Comparaison des valeurs - Heures facturées
    If gValeursAComparer(11, 2) <> gValeursAComparer(11, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(11, 1)), _
                                                "TEC_TdB_Data", CCur(gValeursAComparer(11, 2)), _
                                                "TEC_Local", CCur(gValeursAComparer(11, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If

    'Comparaison des valeurs - Heures TEC
    If gValeursAComparer(12, 2) <> gValeursAComparer(12, 3) Then
        Call ImprimerEcartsVerificationIntegrite(wsOutput, r, CStr(gValeursAComparer(12, 1)), _
                                                "TEC_TdB_Data", CCur(gValeursAComparer(12, 2)), _
                                                "TEC_Local", CCur(gValeursAComparer(12, 3)), _
                                                vbNullString, 0)
        gverificationIntegriteOK = False
    End If

    If gverificationIntegriteOK Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "Tous les montants balancent entre eux")
        r = r + 2
    End If

    'Obtenir le nombre de lignes utilisées dans les tables principales - 2025-01-22 @ 13:46
    Call NoterNombreLignesParFeuille
    
    wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Result print setup - 2024-07-20 @ 14:31
    Dim lastUsedRow As Long
    lastUsedRow = r
    
    'Un peu de couleur
    Set rng = wsOutput.Range("A" & r)
    rng.Value = "**** " & Format$(readRows, "###,##0") & _
                    " lignes analysées dans l'ensemble des tables - " & _
                    Format$(Now(), wsdADMIN.Range("B1").Value & " hh:mm:ss") & " ***"
    rng.Characters(6, 6).Font.Color = vbRed
    rng.Characters(6, 6).Font.Bold = True
    
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "Vérification d'intégrité des tables"
    Dim header2 As String: header2 = vbNullString
    Call MettreEnFormeImpressionSimple(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    If gverificationIntegriteOK = True Then
        MsgBox "La vérification d'intégrité est terminé SANS PROBLÈME" & vbNewLine & vbNewLine & "Voir la feuille 'X_Analyse_Intégrité'", vbInformation
    Else
        MsgBox "La vérification a détecté AU MOINS UN PROBLÈME" & vbNewLine & vbNewLine & "Voir la feuille 'X_Analyse_Intégrité'", vbInformation
    End If
    
    ThisWorkbook.Worksheets("X_Analyse_Integrite").Activate
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set rng = Nothing
    Set rngToPrint = Nothing
    Set wsOutput = Nothing
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierIntegriteTablesLocales", vbNullString, startTime)

End Sub

Sub ImprimerEcartsVerificationIntegrite(ws As Worksheet, r As Long, t As String, _
                                                        s1 As String, v1 As Currency, _
                                                        s2 As String, v2 As Currency, _
                                                        s3 As String, v3 As Currency)

    Dim str1 As String
    Dim str2 As String
    Dim str3 As String

    str1 = Format$(v1, "##,###,##0.00")
    str2 = Format$(v2, "##,###,##0.00")
    str3 = Format$(v3, "##,###,##0.00")
    
    Call AjouterMessageAuxResultats(ws, r, 2, t)
    r = r + 1
    Call AjouterMessageAuxResultats(ws, r, 2, Space(7) & Fn_Pad_A_String(s1, " ", 18, "R") & ":" & Space(14 - Len(str1)) & str1)
    r = r + 1
    Call AjouterMessageAuxResultats(ws, r, 2, Space(7) & Fn_Pad_A_String(s2, " ", 18, "R") & ":" & Space(14 - Len(str2)) & str2)
    r = r + 1
    If Not Trim(s3) = vbNullString Then
        Call AjouterMessageAuxResultats(ws, r, 2, Space(7) & Fn_Pad_A_String(s3, " ", 18, "R") & ":" & Space(14 - Len(str3)) & str3)
        r = r + 1
    End If
    r = r + 1

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

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:CreerOuRemplacerFeuille", vbNullString, 0)
    
    Dim wsExists As Boolean
    wsExists = NomFeuilleExiste(wsName)
    
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

    Call EnregistrerLogApplication("modAppli_Utils:CreerOuRemplacerFeuille", vbNullString, startTime)
    
End Sub

Private Sub VerifierPlanComptable(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierPlanComptable", vbNullString, 0)
    
    Application.ScreenUpdating = True
    
    'dnrPlanComptable_All
    Dim arr As Variant
    Dim nbCol As Long
    nbCol = 4
    arr = Fn_Get_Plan_Comptable(nbCol) 'Returns array with 4 columns (Code, Description)
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(UBound(arr, 1), "###,##0") & _
        " comptes et " & Format$(nbCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de 'dnr_PlanComptable_All'")
    r = r + 1
    
    If UBound(arr, 1) < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim dict_code_GL As New Dictionary
    Dim dict_descr_GL As New Dictionary
    
    Dim i As Long
    Dim codeGL As String
    Dim descrGL As String
    Dim typeGL As String
    Dim cas_doublon_descr As Long
    Dim cas_doublon_code As Long
    Dim cas_type As Long
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        codeGL = arr(i, 1)
        descrGL = arr(i, 2)
        If dict_descr_GL.Exists(descrGL) = False Then
            dict_descr_GL.Add descrGL, codeGL
        Else
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "La description '" & descrGL & "' est un doublon pour le code de G/L '" & codeGL & "'")
            r = r + 1
            cas_doublon_descr = cas_doublon_descr + 1
        End If
        
        If dict_code_GL.Exists(codeGL) = False Then
            dict_code_GL.Add codeGL, descrGL
        Else
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "Le code de G/L '" & codeGL & "' est un doublon pour la description '" & descrGL & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
        
'        GL_ID = arr(i, 3)
        typeGL = arr(i, 4)
        If Len(typeGL) <> 3 Or InStr("A^P^E^R^D^B^I^X^", Left$(typeGL, 1) & "^") = 0 Or IsNumeric(Right$(typeGL, 2)) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "Le type de compte '" & typeGL & "' est INVALIDE pour le code de G/L '" & codeGL & "'")
            r = r + 1
            cas_type = cas_type + 1
        End If
        
    Next i
    
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Un total de " & Format$(UBound(arr, 1), "#,##0") & " comptes ont été analysés!"
    rng.Characters(13, 3).Font.Color = vbRed
    rng.Characters(13, 3).Font.Bold = True
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_descr = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucun doublon de description")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_doublon_descr & " cas de doublons pour les descriptions")
        r = r + 1
    End If
    
    If cas_doublon_code = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucun doublon de code de G/L")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_doublon_code & " cas de doublons pour les codes de G/L")
        r = r + 1
    End If
    
    If cas_type = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucun type de G/L invalide")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_type & " cas de types de G/L invalides")
        r = r + 1
    End If
    r = r + 1
    
    'Cas problème dans cette vérification ?
    If cas_doublon_descr <> 0 Or cas_doublon_descr <> 0 Or cas_type <> 0 Then
        gverificationIntegriteOK = False
    End If
    
Clean_Exit:
    Application.ScreenUpdating = False
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierPlanComptable", vbNullString, startTime)

End Sub

Private Sub VerifierClients(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierClients", vbNullString, 0)
    
    Application.ScreenUpdating = True
    
    Call modImport.ImporterClients
    
    'Fichier maître des Clients
    Dim ws As Worksheet: Set ws = wsdBD_Clients
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(ws.Range("A1").CurrentRegion.Rows.count, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdBD_Clients'")
    r = r + 1
    
    Dim arr As Variant
    arr = wsdBD_Clients.Range("A1").CurrentRegion.Value
    If UBound(arr, 1) < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim dict_code_client As New Dictionary
    Dim dict_nom_client As New Dictionary
    
    Dim i As Long
    Dim code As String
    Dim nom As String
    Dim nomClientSysteme As String
    Dim cas_doublon_nom As Long
    Dim cas_doublon_code As Long
    Dim cas_courriel_invalide As Long
    Dim ligneNonVides As Long
    
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        If Not Trim$(arr(i, 2)) = vbNullString Then
        ligneNonVides = ligneNonVides + 1
            nom = arr(i, 1)
            code = arr(i, 2)
            nomClientSysteme = arr(i, 3)
            
            'Doublon sur le nom ?
            If dict_nom_client.Exists(nom) = False Then
                dict_nom_client.Add nom, code
            Else
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "À la ligne " & i & ", le nom '" & nom & "' est un doublon pour le code '" & code & "'")
                r = r + 1
                cas_doublon_nom = cas_doublon_nom + 1
            End If
            
            'Doublon sur le code de client ?
            If dict_code_client.Exists(code) = False Then
                dict_code_client.Add code, nom
            Else
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "À la ligne " & i & ", le code '" & code & "' est un doublon pour le client '" & nom & "'")
                r = r + 1
                cas_doublon_code = cas_doublon_code + 1
            End If
            
            If Trim$(arr(i, 6)) <> vbNullString Then
                If Fn_ValiderCourriel(arr(i, 6)) = False Then
                    Call AjouterMessageAuxResultats(wsOutput, r, 2, "À la ligne " & i & ", le courriel '" & arr(i, 6) & "' est INVALIDE pour le code '" & code & "'")
                    r = r + 1
                    cas_courriel_invalide = cas_courriel_invalide + 1
                End If
            End If
        End If
    Next i
    
    'Toutes les lignes sont-elles non-vides ?
    If UBound(arr, 1) - 1 <> ligneNonVides Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La feuille comporte un ou des ligne(s) vide(s)")
        r = r + 1
    End If
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Un total de " & Format$(UBound(arr, 1) - 1, "##,##0") & " clients ont été analysés!"
    rng.Characters(13, 5).Font.Color = vbRed
    rng.Characters(13, 5).Font.Bold = True

    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucun doublon de nom")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        r = r + 1
    End If
    
    If cas_doublon_code = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucun doublon de code")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        r = r + 1
    End If
    
    If cas_courriel_invalide = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Toutes les adresses courriel sont valides")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_courriel_invalide & " cas de courriels INVALIDES")
        r = r + 1
    End If
    
    r = r + 1
    
    'Cas problème dans cette vérification ?
    If cas_doublon_nom <> 0 Or _
        cas_doublon_code <> 0 Or _
        cas_courriel_invalide <> 0 Or _
        UBound(arr, 1) - 1 <> ligneNonVides Then
            gverificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set dict_code_client = Nothing
    Set dict_nom_client = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Application.ScreenUpdating = False
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierClients", vbNullString, startTime)

End Sub

Private Sub VerifierFournisseurs(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)
    
    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierFournisseurs", vbNullString, 0)

    Application.ScreenUpdating = True

    Call modImport.ImporterFournisseurs
    
    'wshBD_fournisseurs
    Dim ws As Worksheet: Set ws = wsdBD_Fournisseurs
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(ws.Range("A1").CurrentRegion.Rows.count, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdBD_Fournisseurs'")
    r = r + 1
    
    Dim arr As Variant
    arr = wsdBD_Fournisseurs.Range("A1").CurrentRegion.Value
    If UBound(arr, 1) < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim dict_code_fournisseur As New Dictionary
    Dim dict_nom_fournisseur As New Dictionary
    
    Dim i As Long
    Dim code As String
    Dim nom As String
    Dim cas_doublon_nom As Long
    Dim cas_doublon_code As Long
    
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        nom = arr(i, 1)
        code = arr(i, 2)
        If dict_nom_fournisseur.Exists(nom) = False Then
            dict_nom_fournisseur.Add nom, code
        Else
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "Le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        If dict_code_fournisseur.Exists(code) = False Then
            dict_code_fournisseur.Add code, nom
        Else
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "Le code '" & code & "' est un doublon pour le nom '" & nom & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
    Next i
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Un total de " & Format$(UBound(arr, 1) - 1, "#,##0") & " fournisseurs ont été analysés!"
    rng.Characters(13, 3).Font.Color = vbRed
    rng.Characters(13, 3).Font.Bold = True

    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucun doublon de nom")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        r = r + 1
    End If
    If cas_doublon_code = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucun doublon de code")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        r = r + 1
    End If
    r = r + 1
    
    'Cas problème dans cette vérification ?
    If cas_doublon_nom <> 0 Or cas_doublon_code <> 0 Then
        gverificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set rng = Nothing
    Set ws = Nothing
    
    Application.ScreenUpdating = False
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierFournisseurs", vbNullString, startTime)

End Sub

Private Sub VerifierCCRegularisations(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierCCRegularisations", vbNullString, 0)

    Application.ScreenUpdating = True
    
    Call modImport.ImporterCCRegularisations
    
    'wsdCC_Regularisations
    Dim ws As Worksheet: Set ws = wsdCC_Regularisations
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRowDetails As Long
    lastUsedRowDetails = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRowDetails <= 2 - HeaderRow Or ws.Cells(lastUsedRowDetails, 1).Value = vbNullString Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRowDetails, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    'FAC_Entête Worksheet
    Dim wsFACEntete As Worksheet: Set wsFACEntete = wsdFAC_Entete
    Dim lastUsedRowFacEntete As Long
    lastUsedRowFacEntete = wsFACEntete.Cells(wsFACEntete.Rows.count, 1).End(xlUp).Row
    Dim rngFACEntete As Range: Set rngFACEntete = wsFACEntete.Range("A2:A" & lastUsedRowFacEntete)
    
    'Clients Master File
    Dim wsClients As Worksheet: Set wsClients = wsdBD_Clients
    Dim lastUsedRowClient As Long
    lastUsedRowClient = wsClients.Cells(wsClients.Rows.count, "B").End(xlUp).Row
    Dim rngClients As Range: Set rngClients = wsClients.Range("B2:B" & lastUsedRowClient)
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdCC_Regularisations'")
    r = r + 1
    
    Dim regulNo As Long
    Dim result As Variant
    
    'Dictionary pour accumuler les encaissements par facture
    Dim dictRegul As Scripting.Dictionary
    Set dictRegul = New Scripting.Dictionary
    Dim totalRégularisations As Currency
    
    Dim isRegularisationValid As Boolean
    isRegularisationValid = True
    
    Dim Inv_No As String
    Dim i As Long
    For i = 2 To lastUsedRowDetails
        regulNo = CLng(ws.Cells(i, fREGULRegulID).Value)
        Inv_No = CStr(ws.Cells(i, fREGULInvNo).Value)
        result = Application.WorksheetFunction.XLookup(Inv_No, _
                        rngFACEntete, _
                        rngFACEntete, _
                        "Not Found", _
                        0, _
                        1)
        If result = "Not Found" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La facture '" & Inv_No & "', ligne " & i & ", de la régularisation '" & regulNo & "' n'existe pas dans FAC_Entête")
            r = r + 1
            isRegularisationValid = False
        End If
        
        If IsDate(ws.Cells(i, fREGULDate).Value) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La date '" & ws.Cells(i, fREGULDate).Value & "', ligne " & i & ", de la régularisation '" & regulNo & "' est INVALIDE '")
            r = r + 1
            isRegularisationValid = False
        End If
        
        Dim codeClient As String
        codeClient = ws.Cells(i, fREGULClientID).Value
        result = Application.WorksheetFunction.XLookup(codeClient, _
                        rngClients, _
                        rngClients, _
                        "Not Found", _
                        0, _
                        1)
        If result = "Not Found" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le client '" & codeClient & "' de la régularisation '" & regulNo & "' est INVALIDE")
            r = r + 1
            isRegularisationValid = False
        End If
        
        'Vérification du montant des honoraires
        If IsNumeric(ws.Cells(i, fREGULHono).Value) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le montant des honoraires '" & ws.Cells(i, fREGULHono).Value & "' de la régularisation '" & regulNo & "' n'est pas numérique")
            r = r + 1
            isRegularisationValid = False
        Else
            If dictRegul.Exists(Inv_No) Then
                dictRegul(Inv_No) = dictRegul(Inv_No) + ws.Cells(i, fREGULHono).Value
            Else
                dictRegul.Add Inv_No, ws.Cells(i, fREGULHono).Value
            End If
            totalRégularisations = totalRégularisations + ws.Cells(i, fREGULHono).Value
        End If
        
        'Vérification du montant des frais
        If IsNumeric(ws.Cells(i, fREGULFrais).Value) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le montant des frais '" & ws.Cells(i, fREGULFrais).Value & "' de la régularisation '" & regulNo & "' n'est pas numérique")
            r = r + 1
            isRegularisationValid = False
        Else
            If dictRegul.Exists(Inv_No) Then
                dictRegul(Inv_No) = dictRegul(Inv_No) + ws.Cells(i, fREGULFrais).Value
            Else
                dictRegul.Add Inv_No, ws.Cells(i, fREGULFrais).Value
            End If
            totalRégularisations = totalRégularisations + ws.Cells(i, fREGULFrais).Value
        End If
    
        'Vérification du montant de TPS
        If IsNumeric(ws.Cells(i, fREGULTPS).Value) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le montant de TPS '" & ws.Cells(i, fREGULTPS).Value & "' de la régularisation '" & regulNo & "' n'est pas numérique")
            r = r + 1
            isRegularisationValid = False
        Else
            If dictRegul.Exists(Inv_No) Then
                dictRegul(Inv_No) = dictRegul(Inv_No) + ws.Cells(i, fREGULTPS).Value
            Else
                dictRegul.Add Inv_No, ws.Cells(i, fREGULTPS).Value
            End If
            totalRégularisations = totalRégularisations + ws.Cells(i, fREGULTPS).Value
        End If
        
        'Vérification du montant de TVQ
        If IsNumeric(ws.Cells(i, fREGULTVQ).Value) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le montant de TPS '" & ws.Cells(i, fREGULTVQ).Value & "' de la régularisation '" & regulNo & "' n'est pas numérique")
            r = r + 1
            isRegularisationValid = False
        Else
            If dictRegul.Exists(Inv_No) Then
                dictRegul(Inv_No) = dictRegul(Inv_No) + ws.Cells(i, fREGULTVQ).Value
            Else
                dictRegul.Add Inv_No, ws.Cells(i, fREGULTVQ).Value
            End If
            totalRégularisations = totalRégularisations + ws.Cells(i, fREGULTVQ).Value
        End If
    
    Next i
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(lastUsedRowDetails - 1, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 1
    
    'Compare les régularisations accumulés (dictRegul) avec wsdFAC_Comptes_Clients
    Dim wsComptes_Clients As Worksheet: Set wsComptes_Clients = wsdFAC_Comptes_Clients
    Dim lastUsedRow As Long
    lastUsedRow = wsComptes_Clients.Cells(wsComptes_Clients.Rows.count, 1).End(xlUp).Row
    Dim totalRegul As Currency
    
    For i = 3 To lastUsedRow
        Inv_No = wsComptes_Clients.Cells(i, fFacCCInvNo).Value
        totalRegul = wsComptes_Clients.Cells(i, fFacCCTotalRegul).Value
        If totalRegul <> dictRegul(Inv_No) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Pour la facture '" & Inv_No & "', le total des régularisations de " _
                            & "(wshFAC_Comptes_clients) " & Format$(totalRegul, "###,##0.00 $") _
                            & " est <> du détail des régularisations (wsdCC_Regularisations) " & Format$(dictRegul(Inv_No), "###,##0.00 $"))
            r = r + 1
            isRegularisationValid = False
        End If
    Next i
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Total des régularisations : " & Format$(totalRégularisations, "##,###,##0.00 $")
    rng.Characters(InStr(rng.Value, Left$(totalRégularisations, 1)), 12).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totalRégularisations, 1)), 12).Font.Bold = True
    gValeursAComparer(4, 3) = CCur(totalRégularisations)
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + lastUsedRowDetails - 1
    
    'Cas problème dans cette vérification ?
    If isRegularisationValid = False Then
        gverificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set dictRegul = Nothing
    Set rngClients = Nothing
    Set rngFACEntete = Nothing
    Set ws = Nothing
    Set wsClients = Nothing
    Set wsFACEntete = Nothing
    
    DoEvents
    Application.ScreenUpdating = False
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierCCRegularisations", vbNullString, startTime)

End Sub

Private Sub VerifierDEBRecurrent(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierDEBRecurrent", vbNullString, 0)

    Application.ScreenUpdating = True
    
    Call modImport.ImporterDebRecurrent
    
    Dim ws As Worksheet: Set ws = wsdDEB_Recurrent
    
    'wsdDEB_Recurrent
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= 2 - HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdDEB_Recurrent'")
    r = r + 1
    
    'On a besoin du plan comptable pour valider les données
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wsdADMIN.Range("dnrPlanComptable_All")
    On Error GoTo 0

    Dim isDebRécurrentValid As Boolean
    isDebRécurrentValid = True
    
    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        r = r + 1
        isDebRécurrentValid = False
        GoTo Clean_Exit
    End If
    
    Dim strGL As String
    Dim ligne As Range
    For Each ligne In planComptable.Rows
        strGL = strGL & "^C:" & Trim$(ligne.Cells(1, 2).Value) & "^D:" & Trim$(ligne.Cells(1, 1).Value) & " | "
    Next ligne
    
    'Copie les données vers un tableau
    Dim rng As Range
    Set rng = ws.Range("A1:N" & lastUsedRow)
    Dim arr() As Variant
    Dim headerRows As Long
    headerRows = 1
    Call TransfererPlageVersTableau2D(rng, arr, 1)
    
    'On analyse chacune des lignes du tableau
    Dim i As Long
    Dim p As Long
    Dim GL As String
    Dim descGL As String
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsNumeric(arr(i, 1)) = False Or arr(i, 1) <> Int(arr(i, 1)) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la valeur du numéro de déboursé '" & arr(i, 1) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        
        If IsDate(arr(i, 2)) = False Or arr(i, 2) <> Int(arr(i, 2)) Or arr(i, 2) > Date Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la date '" & arr(i, 2) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        
        p = InStr(strGL, "^C:" & arr(i, 6))
        If p = 0 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le poste de G/L '" & arr(i, 6) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        If p > 0 Then
            GL = Mid$(strGL, p + 3)
            descGL = Mid$(GL, InStr(GL, "^D:") + 3, InStr(GL, " | ") - 8)
            If descGL <> Trim$(arr(i, 7)) Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la description du G/L '" & arr(i, 7) & "' est INVALIDE")
                r = r + 1
                isDebRécurrentValid = False
            End If
        End If
        
        'Total
        If IsNumeric(arr(i, 9)) = False Or arr(i, 9) * 100 <> Int(arr(i, 9) * 100) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant total '" & arr(i, 9) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        'TPS
        If IsNumeric(arr(i, 10)) = False Or arr(i, 10) * 100 <> Int(arr(i, 10) * 100) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant TPS '" & arr(i, 10) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        'TVQ
        If IsNumeric(arr(i, 11)) = False Or arr(i, 11) * 100 <> Int(arr(i, 11) * 100) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant TVQ '" & arr(i, 11) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        'Intrant TPS
        If IsNumeric(arr(i, 12)) = False Or arr(i, 12) * 100 <> Int(arr(i, 12) * 100) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant d'intrant pour la TPS '" & arr(i, 12) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        'Intrant TVQ
        If IsNumeric(arr(i, 13)) = False Or arr(i, 13) * 100 <> Int(arr(i, 13) * 100) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant d'intrant pour la TVQ '" & arr(i, 13) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        '$ dépense
        readRows = readRows + 1
    Next i

    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Cas problème dans cette vérification ?
    If isDebRécurrentValid = False Then
        gverificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set ligne = Nothing
    Set planComptable = Nothing
    Set rng = Nothing
    Set ws = Nothing

    Application.ScreenUpdating = False

    Call EnregistrerLogApplication("modAppli_Utils:VerifierDEBRecurrent", vbNullString, startTime)

End Sub

Private Sub VerifierDEBTrans(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierDEBTrans", vbNullString, 0)

    Application.ScreenUpdating = True
    
    Call modImport.ImporterDebTrans
    
    Dim ws As Worksheet: Set ws = wsdDEB_Trans
    
    'wsdDEB_Trans
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= 2 - HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdDEB_Trans'")
    r = r + 1
    
    'On a besoin du plan comptable pour valider les données
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wsdADMIN.Range("dnrPlanComptable_All")
    On Error GoTo 0

    Dim isDebTransValid As Boolean
    isDebTransValid = True
    
    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        r = r + 1
        isDebTransValid = False
        Exit Sub
    End If
    
    Dim strGL As String
    Dim ligne As Range
    For Each ligne In planComptable.Rows
        strGL = strGL & "^C:" & Trim$(ligne.Cells(1, 2).Value) & "^D:" & Trim$(ligne.Cells(1, 1).Value) & " | "
    Next ligne
    
    'Copie les données vers un tableau
    Dim rng As Range
    Set rng = ws.Range("A1:R" & lastUsedRow)
    Dim arr() As Variant
    Dim headerRows As Long
    headerRows = 1
    Call TransfererPlageVersTableau2D(rng, arr, 1)
    
    'On analyse chacune des lignes du tableau
    Dim i As Long
    Dim p As Long
    Dim GL As String
    Dim descGL As String
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsNumeric(arr(i, 1)) = False Or arr(i, 1) <> Int(arr(i, 1)) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la valeur du numéro de déboursé '" & arr(i, 1) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        
        If IsDate(arr(i, 2)) = False Or arr(i, 2) <> Int(arr(i, 2)) Or arr(i, 2) > Date + 10 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la date '" & arr(i, 2) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        
        p = InStr(strGL, "^C:" & arr(i, 8))
        If p = 0 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le poste de G/L '" & arr(i, 8) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        If p > 0 Then
            GL = Mid$(strGL, p + 3)
            descGL = Mid$(GL, InStr(GL, "^D:") + 3, InStr(GL, " | ") - 8)
            If descGL <> Trim$(arr(i, 9)) Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la description du G/L '" & arr(i, 8) & "' est INVALIDE")
                r = r + 1
                isDebTransValid = False
            End If
        End If
        
        'Total
        If IsNumeric(arr(i, 11)) = False Or arr(i, 11) <> Round(arr(i, 11), 2) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant total '" & arr(i, 11) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        'TPS
        If IsNumeric(arr(i, 12)) = False Or arr(i, 12) <> Round(arr(i, 12), 2) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant TPS '" & arr(i, 12) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        'TVQ
        If IsNumeric(arr(i, 13)) = False Or arr(i, 13) <> Round(arr(i, 13), 2) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant TVQ '" & arr(i, 13) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        'Intrant TPS
        If IsNumeric(arr(i, 14)) = False Or arr(i, 14) <> Round(arr(i, 14), 2) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant d'intrant pour la TPS '" & arr(i, 14) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        'Intrant TVQ
        If IsNumeric(arr(i, 15)) = False Or arr(i, 15) <> Round(arr(i, 15), 2) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant d'intrant pour la TVQ '" & arr(i, 15) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        '$ dépense
        If IsNumeric(arr(i, 16)) = False Or arr(i, 16) <> Round(arr(i, 16), 2) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant de la dépense '" & arr(i, 16) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        readRows = readRows + 1
    Next i

    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Cas problème dans cette vérification ?
    If isDebTransValid = False Then
        gverificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set ligne = Nothing
    Set planComptable = Nothing
    Set rng = Nothing
    Set ws = Nothing
    On Error GoTo 0

    DoEvents
    Application.ScreenUpdating = False

    Call EnregistrerLogApplication("modAppli_Utils:VerifierDEBTrans", vbNullString, startTime)

End Sub

Private Sub VerifierENCDetails(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierENCDetails", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterEncDetails
    
    'wsdENC_Details
    Dim ws As Worksheet: Set ws = wsdENC_Details
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRowDetails As Long
    lastUsedRowDetails = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRowDetails <= 2 - HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRowDetails, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    'ENC_Entête Worksheet
    Dim wsEntete As Worksheet: Set wsEntete = wsdENC_Entete
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsEntete.Cells(wsEntete.Rows.count, 1).End(xlUp).Row
'    Dim rngEntete As Range: Set rngEntete = wsEntete.Range("A2:A" & lastUsedRowEntete)
    Dim strPmtNo As String
    Dim i As Long
    For i = 2 To lastUsedRowEntete
        strPmtNo = strPmtNo & CLng(wsEntete.Cells(i, fEncEPayID).Value) & "|"
    Next i
    
    'FAC_Entête Worksheet
    Dim wsFACEntete As Worksheet: Set wsFACEntete = wsdFAC_Entete
    Dim lastUsedRowFacEntete As Long
    lastUsedRowFacEntete = wsFACEntete.Cells(wsFACEntete.Rows.count, 1).End(xlUp).Row
    Dim rngFACEntete As Range: Set rngFACEntete = wsFACEntete.Range("A2:A" & lastUsedRowFacEntete)
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdENC_Details'")
    r = r + 1
    
    Dim pmtNo As Long, oldpmtNo As Long
    Dim result As Variant
    'Dictionary pour accumuler les encaissements par facture
    Dim dictENC As Scripting.Dictionary
    Set dictENC = New Scripting.Dictionary
    Dim totalEncDetails As Currency
    
    Dim isEncDétailsValid As Boolean
    isEncDétailsValid = True
    
    For i = 2 To lastUsedRowDetails
        pmtNo = CLng(ws.Cells(i, fEncDPayID).Value)
        If pmtNo <> oldpmtNo Then
            If InStr(strPmtNo, pmtNo) = 0 Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le paiement '" & pmtNo & "' à la ligne " & i & " n'existe pas dans ENC_Entête")
                r = r + 1
                isEncDétailsValid = False
            End If
            strPmtNo = strPmtNo & pmtNo & "|"
            oldpmtNo = pmtNo
        End If
        
        Dim Inv_No As String
        Inv_No = CStr(ws.Cells(i, fEncDInvNo).Value)
        result = Application.WorksheetFunction.XLookup(Inv_No, _
                        rngFACEntete, _
                        rngFACEntete, _
                        "Not Found", _
                        0, _
                        1)
        If result = "Not Found" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La facture '" & Inv_No & "', ligne " & i & ", du paiement '" & pmtNo & "' n'existe pas dans FAC_Entête")
            r = r + 1
            isEncDétailsValid = False
        End If
        
        If IsDate(ws.Cells(i, fEncDPayDate).Value) = False Or ws.Cells(i, fEncDPayDate) > Date Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La date '" & ws.Cells(i, fEncDPayDate).Value & "', ligne " & i & ", du paiment '" & pmtNo & "' est INVALIDE '")
            r = r + 1
            isEncDétailsValid = False
        End If
        
        If IsNumeric(ws.Cells(i, fEncDPayAmount).Value) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le montant '" & ws.Cells(i, fEncDPayAmount).Value & "' du paiement '" & pmtNo & "' n'est pas numérique")
            r = r + 1
            isEncDétailsValid = False
        Else
            If dictENC.Exists(Inv_No) Then
                dictENC(Inv_No) = dictENC(Inv_No) + ws.Cells(i, fEncDPayAmount).Value
            Else
                dictENC.Add Inv_No, ws.Cells(i, fEncDPayAmount).Value
            End If
            totalEncDetails = totalEncDetails + ws.Cells(i, fEncDPayAmount).Value
        End If
    Next i
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(lastUsedRowDetails - 1, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 1
    
    'Compare les encaissements accumulés (dictENC) avec wsdFAC_Comptes_Clients
    Dim wsComptes_Clients As Worksheet: Set wsComptes_Clients = wsdFAC_Comptes_Clients
    Dim lastUsedRow As Long
    lastUsedRow = wsComptes_Clients.Cells(wsComptes_Clients.Rows.count, 1).End(xlUp).Row
    Dim totalPaid As Currency
    
    For i = 3 To lastUsedRow
        Inv_No = wsComptes_Clients.Cells(i, fFacCCInvNo).Value
        totalPaid = wsComptes_Clients.Cells(i, fFacCCTotalPaid).Value
        If totalPaid <> dictENC(Inv_No) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Ligne # " & i & " - La facture '" & Inv_No & "', le total des enc. " _
                            & "(wshFAC_Comptes_clients) " & Format$(totalPaid, "###,##0.00 $") _
                            & " est <> du détail des enc. " & Format$(dictENC(Inv_No), "###,##0.00 $"))
            r = r + 1
            isEncDétailsValid = False
        End If
    Next i
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Total des encaissements : " & Format$(totalEncDetails, "##,###,##0.00 $")
    rng.Characters(InStr(rng.Value, Left$(totalEncDetails, 1)), 12).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totalEncDetails, 1)), 12).Font.Bold = True
    gValeursAComparer(3, 4) = CCur(totalEncDetails)
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + lastUsedRowDetails - 1
    
    'Cas problème dans cette vérification ?
    If isEncDétailsValid = False Then
        gverificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set dictENC = Nothing
    Set rng = Nothing
    Set rngFACEntete = Nothing
    Set ws = Nothing
    Set wsComptes_Clients = Nothing
    Set wsFACEntete = Nothing
    Set wsEntete = Nothing
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierENCDetails", vbNullString, startTime)

End Sub

Private Sub VerifierENCEntete(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierENCEntete", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterEncEntete
    
    'Clients Master File
    Dim wsClients As Worksheet: Set wsClients = wsdBD_Clients
    Dim lastUsedRowClient As Long
    lastUsedRowClient = wsClients.Cells(wsClients.Rows.count, "B").End(xlUp).Row
    Dim rngClients As Range: Set rngClients = wsClients.Range("B2:B" & lastUsedRowClient)
    
    'wsdENC_Entete
    Dim ws As Worksheet: Set ws = wsdENC_Entete
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastUsedRow <= HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdENC_Entete'")
    r = r + 1
    
    If lastUsedRow = HeaderRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim arr As Variant
    arr = wsdENC_Entete.Range("A1").CurrentRegion.offset(1, 0) _
              .Resize(lastUsedRow - HeaderRow, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    Dim i As Long
    Dim pmtNo As String
    Dim totals As Currency
    Dim result As Variant
    
    Dim isEncEntêteValid As Boolean
    isEncEntêteValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        pmtNo = arr(i, 1)
        If IsDate(arr(i, 2)) = False Or arr(i, 2) > Date Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La date de paiement '" & arr(i, 2) & "' du paiement '" & arr(i, 1) & "' n'est pas VALIDE")
            r = r + 1
            isEncEntêteValid = False
        End If
        
        Dim codeClient As String
        codeClient = arr(i, 4)
        result = Application.WorksheetFunction.XLookup(codeClient, _
                        rngClients, _
                        rngClients, _
                        "Not Found", _
                        0, _
                        1)
        If result = "Not Found" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le client '" & codeClient & "' du paiement '" & pmtNo & "' est INVALIDE")
            r = r + 1
            isEncEntêteValid = False
        End If
        totals = totals + arr(i, 6)
    Next i
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
    r = r + 1
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Total des encaissements : " & Format$(totals, "##,###,##0.00 $")
    rng.Characters(InStr(rng.Value, Left$(totals, 1)), 12).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totals, 1)), 12).Font.Bold = True
    gValeursAComparer(3, 3) = CCur(totals)
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    'Cas problème dans cette vérification ?
    If isEncEntêteValid = False Then
        gverificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set rng = Nothing
    Set rngClients = Nothing
    Set ws = Nothing
    Set wsClients = Nothing
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierENCEntete", vbNullString, startTime)

End Sub

Private Sub VerifierFACDetails(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierFACDetails", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterFacDetails
    
    'wsdFAC_Details
    Dim ws As Worksheet: Set ws = wsdFAC_Details
    Dim HeaderRow As Long: HeaderRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wsdFAC_Entete
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A" & 1 + HeaderRow & ":A" & lastUsedRowEntete)
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdFAC_Details'")
    r = r + 1
    
    'Transfer FAC_Details data from Worksheet into an Array (arr)
    Dim arr As Variant
    arr = wsdFAC_Details.Range("A1").CurrentRegion.offset(1, 0).Value
    
    Dim i As Long
    Dim Inv_No As String, oldInv_No As String
    Dim result As Variant
    
    Dim isFACDétailsValid As Boolean
    isFACDétailsValid = True
    
    For i = LBound(arr, 1) + 2 To UBound(arr, 1) - 1 'Two lines of header !
        Inv_No = CStr(arr(i, 1))
'        Debug.Print "#018 - Inv_no = ", Inv_No, ", de type ", TypeName(Inv_No)
        If Inv_No <> oldInv_No Then
             result = Application.WorksheetFunction.XLookup(Inv_No, _
                                                    rngMaster, _
                                                    rngMaster, _
                                                    "Not Found", _
                                                    0, _
                                                    1)
            If result = "Not Found" Then
                Debug.Print "#019 " & result
            End If
            oldInv_No = CStr(Inv_No)
        End If
        If result = "Not Found" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " n'existe pas dans FAC_Entête")
            r = r + 1
            isFACDétailsValid = False
        End If
        If IsNumeric(arr(i, 3)) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " le nombre d'heures est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
            isFACDétailsValid = False
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " le taux horaire est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
            isFACDétailsValid = False
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " le montant est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
            isFACDétailsValid = False
        End If
    Next i
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 2, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - 2
    
    'Cas problème dans cette vérification ?
    If isFACDétailsValid = False Then
        gverificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set rngMaster = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierFACDetails", vbNullString, startTime)

End Sub

Private Sub VerifierFACEntete(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierFACEntete", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterFacEntete
    
    'wsdFAC_Entete
    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    
    Dim HeaderRow As Long: HeaderRow = 2
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdFAC_Entete'")
    r = r + 1
    
    If lastUsedRow = HeaderRow Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    'Prépare à charger les données en mémoire (arr)
    Dim rngData As Range
    Set rngData = ws.Range("A1").CurrentRegion
    Set rngData = rngData.offset(2, 0).Resize(rngData.Rows.count - 2, rngData.Columns.count)
    Dim arr As Variant
    arr = rngData
    
    Dim i As Long
    Dim Inv_No As String
    Dim totals(1 To 8, 1 To 2) As Currency
    Dim nbFactC As Long, nbFactAC As Long
    
    Dim isFACEntêteValid As Boolean
    isFACEntêteValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        Inv_No = arr(i, 1)
        If IsDate(arr(i, 2)) = False Or arr(i, 2) > Date Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
            isFACEntêteValid = False
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & ", la date est de mauvais format '" & arr(i, 2) & "'")
                r = r + 1
                isFACEntêteValid = False
            End If
        End If
        If arr(i, 3) <> "C" And arr(i, 3) <> "AC" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le type de facture '" & arr(i, 3) & "' pour la facture '" & Inv_No & "' est INVALIDE")
            isFACEntêteValid = False
            r = r + 1
        End If
        If arr(i, 19) <> 0.09975 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le % de TVQ, soit '" & arr(i, 19) & "' pour la facture '" & Inv_No & "' est ERRONÉ")
            r = r + 1
            isFACEntêteValid = False
        End If
        If arr(i, 3) = "C" Then
            totals(1, 1) = totals(1, 1) + arr(i, fFacEHonoraires)
            totals(2, 1) = totals(2, 1) + arr(i, fFacEAutresFrais1)
            totals(3, 1) = totals(3, 1) + arr(i, fFacEAutresFrais2)
            totals(4, 1) = totals(4, 1) + arr(i, fFacEAutresFrais3)
            totals(5, 1) = totals(5, 1) + arr(i, fFacEMntTPS)
            totals(6, 1) = totals(6, 1) + arr(i, fFacEMntTVQ)
            totals(7, 1) = totals(7, 1) + arr(i, fFacEARTotal)
            totals(8, 1) = totals(8, 1) + arr(i, fFacEDépôt)
            nbFactC = nbFactC + 1
        Else
            totals(1, 2) = totals(1, 2) + arr(i, fFacEHonoraires)
            totals(2, 2) = totals(2, 2) + arr(i, fFacEAutresFrais1)
            totals(3, 2) = totals(3, 2) + arr(i, fFacEAutresFrais2)
            totals(4, 2) = totals(4, 2) + arr(i, fFacEAutresFrais3)
            totals(5, 2) = totals(5, 2) + arr(i, fFacEMntTPS)
            totals(6, 2) = totals(6, 2) + arr(i, fFacEMntTVQ)
            totals(7, 2) = totals(7, 2) + arr(i, fFacEARTotal)
            totals(8, 2) = totals(8, 2) + arr(i, fFacEDépôt)
            nbFactAC = nbFactAC + 1
        End If
    Next i
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
    r = r + 1
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Totaux des factures CONFIRMÉES (" & nbFactC & " factures)"
    rng.Characters(InStr(rng.Value, "CONFIRMÉES"), 10).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, "CONFIRMÉES"), 10).Font.Bold = True
    r = r + 1

    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Total Fact. : " & Fn_Pad_A_String(Format$(totals(7, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left$(totals(7, 1), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totals(7, 1), 1)), 15).Font.Bold = True
    gValeursAComparer(1, 1) = "Total Factures (Confirmées)"
    gValeursAComparer(1, 2) = CCur(totals(7, 1))
    r = r + 1

    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Acompte payé: " & _
            Fn_Pad_A_String(Format$(totals(8, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 2
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "Totaux des factures À CONFIRMER (" & nbFactAC & " factures)"
    rng.Characters(InStr(rng.Value, "À CONFIRMER"), 11).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, "À CONFIRMER"), 11).Font.Bold = True
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Total Fact. : " & Fn_Pad_A_String(Format$(totals(7, 2), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left$(totals(7, 2), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totals(7, 2), 1)), 15).Font.Bold = True
    gValeursAComparer(2, 1) = "Total Factures (À confirmer)"
    gValeursAComparer(2, 2) = CCur(totals(7, 2))
    r = r + 1

    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Acompte payé: " & _
            Fn_Pad_A_String(Format$(totals(8, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - HeaderRow
    
    'Cas problème dans cette vérification ?
    If isFACEntêteValid = False Then
        gverificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set rng = Nothing
    Set rngData = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierFACEntete", vbNullString, startTime)

End Sub

Private Sub VerifierFACComptesClients(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierFACComptesClients", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterFacComptesClients
    
    'wsdGL_Trans
    Dim ws As Worksheet: Set ws = wsdFAC_Comptes_Clients
    
    Dim HeaderRow As Long: HeaderRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdFAC_Comptes_Clients'")
    r = r + 1
    
    If lastUsedRow = HeaderRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    'Load every records into an Array
    Dim arr As Variant
    arr = wsdFAC_Comptes_Clients.Range("A1").CurrentRegion.offset(2, 0) _
              .Resize(lastUsedRow - HeaderRow, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    Dim i As Long
    Dim Inv_No As String
    Dim totals(1 To 4, 1 To 2) As Currency
    Dim nbFactC As Long, nbFactAC As Long
    
    Dim isFACCCValid As Boolean
    isFACCCValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        Inv_No = arr(i, fFacCCInvNo)
        Dim invType As String
        invType = Fn_Get_Invoice_Type(Inv_No)
        If invType <> "C" And invType <> "AC" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le type de facture '" & invType & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Date ?
        If IsDate(CDate(arr(i, fFacCCInvoiceDate))) = False Or arr(i, fFacCCInvoiceDate) > Date Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", la date '" & arr(i, fFacCCInvoiceDate) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        Else
            If arr(i, fFacCCInvoiceDate) <> Int(arr(i, fFacCCInvoiceDate)) Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", la facture '" & Inv_No & "', la date est de mauvais format '" & arr(i, fFacCCInvoiceDate) & "'")
                r = r + 1
                isFACCCValid = False
            End If
        End If
        
        'Code client ?
        If Fn_Validate_Client_Number(CStr(arr(i, fFacCCCodeClient))) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le client '" & CStr(arr(i, fFacCCCodeClient)) & "' de la facture '" & Inv_No & "' est INVALIDE '")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Status (Paid or Unpaid)
        If arr(i, fFacCCStatus) <> "Paid" And arr(i, fFacCCStatus) <> "Unpaid" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le statut '" & arr(i, fFacCCStatus) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Date due
        If IsDate(CDate(arr(i, fFacCCDueDate))) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", la date due '" & arr(i, fFacCCDueDate) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Total
        If IsNumeric(arr(i, fFacCCTotal)) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le total de la facture '" & arr(i, fFacCCTotal) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Encaissé à date
        If IsNumeric(arr(i, fFacCCTotalPaid)) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le montant payé à date '" & arr(i, fFacCCTotalPaid) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Régularisation à date
        If IsNumeric(arr(i, fFacCCTotalRegul)) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le montant régularisé à date '" & arr(i, fFacCCTotalRegul) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Solde à recevoir
        If IsNumeric(arr(i, fFacCCBalance)) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le solde de la facture '" & arr(i, fFacCCBalance) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'PLUG pour s'assurer que le solde impayé est bel et bien aligné sur le total moins le payé à date + les régularisations
        If Round(arr(i, fFacCCBalance), 2) <> Round(arr(i, fFacCCTotal) - arr(i, fFacCCTotalPaid) + arr(i, fFacCCTotalRegul), 2) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", pour la facture '" & Inv_No & _
                    ", le solde à recevoir est erroné, soit " & Format$(CCur(arr(i, fFacCCBalance)), "###,##0.00 $") & "' vs. " & _
                    "'" & Format$(arr(i, fFacCCTotal) - arr(i, fFacCCTotalPaid) + arr(i, fFacCCTotalRegul), "###,##0.00 $") & "'")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Le statut est-il exact
        If arr(i, fFacCCBalance) = 0 And arr(i, fFacCCStatus) = "Unpaid" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le statut '" & arr(i, fFacCCStatus) & "' de la facture '" & Inv_No & "', avec un solde de " & Format$(arr(i, fFacCCBalance), "#,##0.00 $") & " est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        If arr(i, fFacCCBalance) <> 0 And arr(i, fFacCCStatus) = "Paid" Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le statut '" & arr(i, fFacCCStatus) & "' de la facture '" & Inv_No & "', avec un solde de " & Format$(arr(i, fFacCCBalance), "#,##0.00 $") & " est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        If invType = "C" Then
            totals(1, 1) = totals(1, 1) + arr(i, fFacCCTotal)
            totals(2, 1) = totals(2, 1) + arr(i, fFacCCTotalPaid)
            totals(3, 1) = totals(3, 1) + arr(i, fFacCCTotalRegul)
            totals(4, 1) = totals(4, 1) + arr(i, fFacCCBalance)
            nbFactC = nbFactC + 1
        Else
            totals(1, 2) = totals(1, 2) + arr(i, fFacCCTotal)
            totals(2, 2) = totals(2, 2) + arr(i, fFacCCTotalPaid)
            totals(3, 2) = totals(3, 2) + arr(i, fFacCCTotalRegul)
           totals(4, 2) = totals(4, 2) + arr(i, fFacCCBalance)
            nbFactAC = nbFactAC + 1
        End If
    Next i
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
    r = r + 1
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Totaux des factures CONFIRMÉES (" & nbFactC & " factures)"
    rng.Characters(InStr(rng.Value, "CONFIRMÉES"), 10).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, "CONFIRMÉES"), 10).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Total des factures         : " & Fn_Pad_A_String(Format$(totals(1, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left$(totals(1, 1), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totals(1, 1), 1)), 15).Font.Bold = True
    gValeursAComparer(1, 3) = CCur(totals(1, 1))
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Montants encaissés à date  : " & Fn_Pad_A_String(Format$(totals(2, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left$(totals(2, 1), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totals(2, 1), 1)), 15).Font.Bold = True
    gValeursAComparer(3, 1) = "Total des montants encaissés à date"
    gValeursAComparer(3, 2) = CCur(totals(2, 1))
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Montant régularisé à date  : " & Fn_Pad_A_String(Format$(totals(3, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left$(totals(3, 1), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totals(3, 1), 1)), 15).Font.Bold = True
    gValeursAComparer(4, 1) = "Total régularisé à date"
    gValeursAComparer(4, 2) = CCur(totals(3, 1))
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Solde à recevoir           : " & Fn_Pad_A_String(Format$(totals(4, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left$(totals(4, 1), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totals(4, 1), 1)), 15).Font.Bold = True
    gValeursAComparer(5, 1) = "Solde à recevoir"
    gValeursAComparer(5, 2) = CCur(totals(4, 1))
    r = r + 2
    gsoldeComptesClients = totals(4, 1)
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "Total des factures À CONFIRMER (" & nbFactAC & " factures)"
    rng.Characters(InStr(rng.Value, "À CONFIRMER"), 11).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, "À CONFIRMER"), 11).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Total des factures        : " & Fn_Pad_A_String(Format$(totals(1, 2), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left$(totals(1, 2), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(totals(1, 2), 1)), 15).Font.Bold = True
    gValeursAComparer(2, 3) = CCur(totals(1, 2))
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - HeaderRow
    
    'Cas problème dans cette vérification ?
    If isFACCCValid = False Then
        gverificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set rng = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierFACComptesClients", vbNullString, startTime)

End Sub

Private Sub VerifierFACSommaireTaux(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierFACSommaireTaux", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterFacSommaireTaux
    
    Dim ws As Worksheet: Set ws = wsdFAC_Sommaire_Taux
    
    Dim tol As Double
    tol = 0.0001 'Petite tolérance pour la comparaison
    
    'wsdFAC_Sommaire_Taux
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= 2 - HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdFAC_Sommaire_Taux'")
    r = r + 1
    
    'On a besoin des factures pour la validation
    Dim wsMaster As Worksheet: Set wsMaster = wsdFAC_Entete
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A2:A" & lastUsedRowEntete)
    
    'On a besoin des professionnels
    Dim rngProf As Range
    Call Get_Range_From_Dynamic_Named_Range("dnrProf_Initials_Only", rngProf)

    'Copie les données vers un tableau
    Dim rng As Range
    Set rng = ws.Range("A1:E" & lastUsedRow)
    Dim arr() As Variant
    
    Dim isFACSTValid As Boolean
    isFACSTValid = True
    
    Dim headerRows As Long
    headerRows = 1
    Call TransfererPlageVersTableau2D(rng, arr, 1)
    
    'On analyse chacune des lignes du tableau
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Fn_Is_String_Valid(CStr(arr(i, 1)), rngMaster) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la facture '" & arr(i, 1) & "' est INVALIDE")
            r = r + 1
            isFACSTValid = False
        End If
        If IsNumeric(arr(i, 2)) = False Or arr(i, 2) <> CLng(Int(arr(i, 2))) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la valeur de la séquence '" & arr(i, 2) & "' est INVALIDE")
            r = r + 1
            isFACSTValid = False
        End If
        
        If Fn_Is_String_Valid(CStr(arr(i, 3)), rngProf) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le professionnel '" & arr(i, 3) & "' est INVALIDE")
            r = r + 1
            isFACSTValid = False
        End If
        
        'Heures
        If IsNumeric(arr(i, 4)) = False Or Abs((arr(i, 4) * 100) - Round(arr(i, 4) * 100, 0)) > tol Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", les heures '" & arr(i, 4) & "' sont INVALIDES")
            r = r + 1
            isFACSTValid = False
        End If
        
        'Taux Horaire
        If IsNumeric(arr(i, 5)) = False Or Abs((arr(i, 5) * 100) - Round(arr(i, 5) * 100, 0)) > tol Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le taux horaire '" & arr(i, 5) & "' est INVALIDE")
            r = r + 1
            isFACSTValid = False
        End If
        readRows = readRows + 1
    Next i

    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Cas problème dans cette vérification ?
    If isFACSTValid = False Then
        gverificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set rng = Nothing
    Set rngMaster = Nothing
    Set rngProf = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    On Error GoTo 0

    Application.ScreenUpdating = True

    Call EnregistrerLogApplication("modAppli_Utils:VerifierFACSommaireTaux", vbNullString, startTime)

End Sub

Private Sub VerifierFACProjetsEntete(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierFACProjetsEntete", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterFacProjetsEntete
    
    'wsdGL_Trans
    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Entete
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= HeaderRow Or ws.Cells(2, 1).Value = vbNullString Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdFAC_Projets_Entete'")
    r = r + 1
    
    'Establish the number of rows before transferring it to an Array
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.offset(1, 0).Resize(lastUsedRow - 1, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    Dim i As Long
    Dim projetID As String
    Dim codeClient As String
    
    Dim isFacProjetEntêteValid As Boolean
    isFacProjetEntêteValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) <> vbNullString Then
            projetID = arr(i, 1)
            'Client valide ?
            codeClient = Trim$(arr(i, 3))
            If Fn_Validate_Client_Number(codeClient) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le Code de Client est INVALIDE '" & arr(i, 3) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If IsDate(arr(i, 4)) = False Or arr(i, 4) > Date Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 4) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If IsNumeric(arr(i, 5)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le total des honoraires est INVALIDE '" & arr(i, 5) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If IsNumeric(arr(i, 7)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du premier sommaire sont INVALIDES '" & arr(i, 7) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If IsNumeric(arr(i, 8)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du premier sommaire est INVALIDE '" & arr(i, 8) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If IsNumeric(arr(i, 9)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du premier sommaire sont INVALIDES '" & arr(i, 9) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 11) <> vbNullString And IsNumeric(arr(i, 11)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du second sommaire sont INVALIDES '" & arr(i, 11) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 12) <> vbNullString And IsNumeric(arr(i, 12)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du second sommaire est INVALIDE '" & arr(i, 12) & "'")
                r = r + 1
            End If
            If arr(i, 13) <> vbNullString And IsNumeric(arr(i, 13)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du second sommaire sont INVALIDES '" & arr(i, 13) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 15) <> vbNullString And IsNumeric(arr(i, 15)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du troisième sommaire sont INVALIDES '" & arr(i, 15) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 16) <> vbNullString And IsNumeric(arr(i, 16)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du troisième sommaire est INVALIDE '" & arr(i, 16) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 17) <> vbNullString And IsNumeric(arr(i, 17)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du troisième sommaire sont INVALIDES '" & arr(i, 17) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 19) <> vbNullString And IsNumeric(arr(i, 19)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du quatrième sommaire sont INVALIDES '" & arr(i, 19) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 20) <> vbNullString And IsNumeric(arr(i, 20)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du quatrième sommaire est INVALIDE '" & arr(i, 20) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 21) <> vbNullString And IsNumeric(arr(i, 21)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du quatrième sommaire sont INVALIDES '" & arr(i, 21) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 23) <> vbNullString And IsNumeric(arr(i, 23)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du cinquième sommaire sont INVALIDES '" & arr(i, 23) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 24) <> vbNullString And IsNumeric(arr(i, 24)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du cinquième sommaire est INVALIDE '" & arr(i, 24) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
            If arr(i, 25) <> vbNullString And IsNumeric(arr(i, 25)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du cinquième sommaire sont INVALIDES '" & arr(i, 25) & "'")
                r = r + 1
                isFacProjetEntêteValid = False
            End If
        End If
    Next i
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " projets de factures a été analysés")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    'Cas problème dans cette vérification ?
    If isFacProjetEntêteValid = False Then
        gverificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set ws = Nothing
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierFACProjetsEntete", vbNullString, startTime)

End Sub

Private Sub VerifierFACProjetsDetails(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierFACProjetsDetails", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterFacProjetsDetails
    
    'wsdFAC_Projets_Details
    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Details
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= HeaderRow Or ws.Cells(2, 1).Value = vbNullString Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If

    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim lastUsedRowEntete As Long
    Dim wsMaster As Worksheet: Set wsMaster = wsdFAC_Projets_Entete
    lastUsedRowEntete = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A2:A" & lastUsedRowEntete)
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdFAC_Projets_Details'")
    r = r + 1
    
    'Charge le contenu de 'wsdFAC_Projets_Details' en mémoire (Array)
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.offset(1, 0).Resize(lastUsedRow - 1, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    Dim i As Long
    Dim projetID As Long, oldProjetID As Long
    Dim codeClient As String
    Dim result As Variant
    
    Dim isFacProjetDetailValid As Boolean
    isFacProjetDetailValid = True
    
    'Sauvegarde la ligne active
    Dim saveR As Long
    saveR = r
    
    'À partir de la mémoire (Array)
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) <> vbNullString Then
            projetID = CLng(arr(i, 1))
            If projetID <> oldProjetID Then
                result = Application.WorksheetFunction.XLookup(projetID, _
                                     rngMaster, _
                                     rngMaster, _
                                     "Not Found", _
                                     0, _
                                     1)
                oldProjetID = projetID
            End If
            If result = "Not Found" Then
                 Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le projet '" & projetID & "' à la ligne " & i & " n'existe pas dans FAC_Projets_Entête")
                 r = r + 1
                 isFacProjetDetailValid = False
            End If
            'Client valide ?
            codeClient = Trim$(arr(i, 3))
            If Fn_Validate_Client_Number(codeClient) = False Then
                 Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le Code de Client est INVALIDE '" & arr(i, 3) & "'")
                 r = r + 1
                 isFacProjetDetailValid = False
            End If
            If IsNumeric(arr(i, 4)) = False Then
                 Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le projet '" & projetID & "' à la ligne " & i & " le TECID est INVALIDE '" & arr(i, 4) & "'")
                 r = r + 1
            End If
            If IsNumeric(arr(i, 5)) = False Then
                 Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le projet '" & projetID & "' à la ligne " & i & " le ProfID est INVALIDE '" & arr(i, 5) & "'")
                 r = r + 1
                 isFacProjetDetailValid = False
            End If
            If IsDate(arr(i, 6)) = False Or CDate(arr(i, 6)) > Date Then
                 Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le projet '" & projetID & "' à la ligne " & i & " la Date est INVALIDE '" & arr(i, 6) & "'")
                 r = r + 1
                 isFacProjetDetailValid = False
            End If
            If IsNumeric(arr(i, 8)) = False Then
                 Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le projet '" & projetID & "' à la ligne " & i & " les Heures sont INVALIDES '" & arr(i, 8) & "'")
                 r = r + 1
                 isFacProjetDetailValid = False
            End If
        End If
    Next i
    
    'Est-ce qu'il y a eu des messages de générés ?
    If saveR <> r Then
        gverificationIntegriteOK = False
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes ont été analysées")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - HeaderRow
    
    'Cas problème dans cette vérification ?
    If isFacProjetDetailValid = False Then
        gverificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set rngMaster = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierFACProjetsDetails", vbNullString, startTime)

End Sub

Private Sub VerifierGLTrans(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierGLTrans", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterGLTransactions
    
    'wsdGL_Trans
    Dim ws As Worksheet: Set ws = wsdGL_Trans
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastUsedRow <= HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim firstEmptyCol As Long
    firstEmptyCol = 1
    Do Until ws.Cells(HeaderRow, firstEmptyCol) = vbNullString
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdGL_Trans'")
    r = r + 1
    
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wsdADMIN.Range("dnrPlanComptable_All")
    On Error GoTo 0

    Dim isGLTransValid As Boolean
    isGLTransValid = True
    
    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        r = r + 1
        isGLTransValid = False
        Exit Sub
    End If
    
    Dim strCodeGL As String, strDescGL As String
    Dim ligne As Range
    For Each ligne In planComptable.Rows
        strCodeGL = strCodeGL & ligne.Cells(1, 2).Value & "|:|"
        strDescGL = strDescGL & ligne.Cells(1, 1).Value & "|:|"
    Next ligne
    
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.Rows.count - 1 'Remove the header row
    If numRows < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.offset(1, 0).Resize(numRows, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    Dim dict_GL_Entry As New Dictionary
    Dim sum_arr() As Currency
    ReDim sum_arr(1 To 2500, 1 To 3)
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim dt As Currency, ct As Currency
    Dim arTotal As Currency
    Dim GL_Entry_No As String, glCode As String, glDescr As String
    Dim CCGlNo As String
    CCGlNo = ObtenirNoGlIndicateur("Comptes Clients")
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        GL_Entry_No = arr(i, 1)
        If dict_GL_Entry.Exists(GL_Entry_No) = False Then
            dict_GL_Entry.Add GL_Entry_No, row
            sum_arr(row, 1) = GL_Entry_No
            row = row + 1
        End If
        If IsDate(arr(i, 2)) = False Or arr(i, 2) > Date + 10 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** L'écriture #  " & GL_Entry_No & " ' à la ligne " & i & " a une date INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
            isGLTransValid = False
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** L'écriture #  " & GL_Entry_No & " ' à la ligne " & i & " a une date avec le mauvais format '" & arr(i, 2) & "'")
                r = r + 1
                isGLTransValid = False
            End If
        End If
        glCode = CStr(arr(i, 5))
        If InStr(1, strCodeGL, glCode + "|:|") = 0 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le compte '" & glCode & "' à la ligne " & i & " est INVALIDE '")
            r = r + 1
            isGLTransValid = False
        End If
        If glCode = CCGlNo Then
            arTotal = arTotal + arr(i, 7) - arr(i, 8)
        End If
        glDescr = arr(i, 6)
        If InStr(1, strDescGL, glDescr + "|:|") = 0 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La description du compte '" & glDescr & "' à la ligne " & i & " est INVALIDE")
            r = r + 1
            isGLTransValid = False
        End If
        dt = arr(i, 7)
        If IsNumeric(dt) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le montant du débit '" & dt & "' à la ligne " & i & " n'est pas une valeur numérique")
            r = r + 1
            isGLTransValid = False
        End If
        ct = arr(i, 8)
        If IsNumeric(ct) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le montant du débit '" & ct & "' à la ligne " & i & " n'est pas une valeur numérique")
            r = r + 1
            isGLTransValid = False
        End If
        currentRow = dict_GL_Entry(GL_Entry_No)
        sum_arr(currentRow, 2) = sum_arr(currentRow, 2) + dt
        sum_arr(currentRow, 3) = sum_arr(currentRow, 3) + ct
        If arr(i, 10) <> vbNullString Then
            If IsDate(arr(i, 10)) = False Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le TimeStamp '" & arr(i, 10) & "' à la ligne " & i & " n'est pas une date VALIDE")
                r = r + 1
                isGLTransValid = False
            End If
        End If
    Next i
    
    Dim sum_dt As Currency, sum_ct As Currency
    Dim cas_hors_balance As Long
    Dim v As Variant
    For Each v In dict_GL_Entry.items()
        GL_Entry_No = sum_arr(v, 1)
        dt = Round(sum_arr(v, 2), 2)
        ct = Round(sum_arr(v, 3), 2)
        If dt <> ct Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Écriture # " & GL_Entry_No & " ne balance pas... Dt = " & Format$(dt, "###,###,##0.00") & " et Ct = " & Format$(ct, "###,###,##0.00"))
            r = r + 1
            isGLTransValid = False
            cas_hors_balance = cas_hors_balance + 1
        End If
        sum_dt = sum_dt + dt
        sum_ct = sum_ct + ct
    Next v
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - HeaderRow, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - HeaderRow
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & dict_GL_Entry.count & " écritures ont été analysées")
    r = r + 1
    
    If cas_hors_balance = 0 Then
        'Un peu de couleur
        Dim rng As Range: Set rng = wsOutput.Range("B" & r)
        rng.Value = "       Chacune des écritures balancent au niveau de l'écriture"
        rng.Characters(InStr(rng.Value, "C"), Len(rng.Value) - 7).Font.Color = vbRed
        rng.Characters(InStr(rng.Value, "C"), Len(rng.Value) - 7).Font.Bold = True
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_hors_balance & " écriture(s) qui ne balance(nt) pas !!!")
        r = r + 1
        isGLTransValid = False
    End If
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Les totaux des transactions sont:")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Dt = " & Format$(sum_dt, "###,###,##0.00 $"))
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Ct = " & Format$(sum_ct, "###,###,##0.00 $"))
    r = r + 1
    
    If sum_dt - sum_ct <> 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Hors-Balance de " & Format$(sum_dt - sum_ct, "###,###,##0.00 $"))
        r = r + 1
        isGLTransValid = False
    End If
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "Au Grand Livre, le solde des Comptes-Clients est de : " & Format$(arTotal, "##,###,##0.00 $")
    rng.Characters(InStr(rng.Value, Left$(arTotal, 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, Left$(arTotal, 1)), 15).Font.Bold = True
    gValeursAComparer(5, 3) = CCur(arTotal)
    r = r + 2
    If gsoldeComptesClients <> arTotal Then
        MsgBox "ATTENTION, le solde des Comptes-Clients" & vbNewLine & vbNewLine & _
                "diffère entre les 2 sources...", vbCritical, "FAC_Comptes_Clients <> Solde au Grand-Livre !!!"
    End If
    
    'Cas problème dans cette vérification ?
    If isGLTransValid = False Then
        gverificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set ligne = Nothing
    Set planComptable = Nothing
    Set rng = Nothing
    Set v = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierGLTrans", vbNullString, startTime)

End Sub

Private Sub VerifierGLEJRecurrente(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierGLEJRecurrente", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterEJRecurrente
    
    Dim ws As Worksheet: Set ws = wsdGL_EJ_Recurrente
    
    'wsdGL_EJ_Recurrente
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= 2 - HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdGL_EJ_Recurrente'")
    r = r + 1
    
    Dim isGlEjRécurrenteValid As Boolean
    isGlEjRécurrenteValid = True
    
    'On a besoin des comptes du G/L pour la validation
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wsdADMIN.Range("dnrPlanComptable_All")
    On Error GoTo 0

    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation, "modAppli_Utils:VerifierGLEJRecurrente"
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        r = r + 1
        isGlEjRécurrenteValid = False
        Exit Sub
    End If
    
    'Bâtir une chaine avec code & description
    Dim strGL As String
    Dim ligne As Range
    For Each ligne In planComptable.Rows
        strGL = strGL & Trim$(ligne.Cells(1, 2).Value) & "-" & Trim$(ligne.Cells(1, 1).Value) & " | "
    Next ligne

    'Copier les données vers un tableau
    Dim rng As Range
    Set rng = ws.Range("A1:G" & lastUsedRow)
    Dim arr() As Variant
    Dim headerRows As Long
    headerRows = 1
    Call TransfererPlageVersTableau2D(rng, arr, 1)
    
    'On analyse chacune des lignes du tableau
    Dim i As Long, p As Long
    Dim GL As String, descGL As String
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsNumeric(arr(i, 1)) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le numéro d'écriture '" & arr(i, 1) & "' est INVALIDE")
            r = r + 1
            isGlEjRécurrenteValid = False
        End If
        
        p = InStr(strGL, Trim$(arr(i, 3)))
        If p = 0 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le poste de G/L '" & arr(i, 3) & "' est INVALIDE")
            r = r + 1
            isGlEjRécurrenteValid = False
        End If
        If p > 0 Then
            GL = Mid$(strGL, p)
            descGL = Mid$(GL, InStr(GL, "-") + 1, InStr(GL, " | ") - 6)
            If descGL <> Trim$(arr(i, 4)) Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la description du G/L '" & arr(i, 4) & "' est INVALIDE")
                r = r + 1
                isGlEjRécurrenteValid = False
            End If
        End If
        If arr(i, 5) <> vbNullString Then
            If IsNumeric(arr(i, 5)) = False Or arr(i, 5) <> Round(arr(i, 5), 2) Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le $ du débit '" & arr(i, 5) & "' est INVALIDE")
                r = r + 1
                isGlEjRécurrenteValid = False
            End If
        End If
        If arr(i, 6) <> vbNullString Then
            If IsNumeric(arr(i, 6)) = False Or arr(i, 6) <> Round(arr(i, 6), 2) Then
                Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le $ du crédit '" & arr(i, 6) & "' est INVALIDE")
                r = r + 1
                isGlEjRécurrenteValid = False
            End If
        End If
        readRows = readRows + 1
    Next i

    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Cas problème dans cette vérification ?
    If isGlEjRécurrenteValid = False Then
        gverificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set ligne = Nothing
    Set planComptable = Nothing
    Set rng = Nothing
    Set ws = Nothing
    On Error GoTo 0

    Application.ScreenUpdating = True

    Call EnregistrerLogApplication("modAppli_Utils:VerifierGLEJRecurrente", vbNullString, startTime)

End Sub

Private Sub VerifierTECTdBData(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierTECTdBData", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Call modImport.ImporterTEC
    Call ActualiserTECTableauDeBord
    
    'wshTEC_TdB_Data
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastUsedRow <= HeaderRow Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshTEC_TdB_Data'")
    r = r + 1
    
    'On a besoin des professionnels
    Dim rngProf As Range
    Call Get_Range_From_Dynamic_Named_Range("dnrProf_Initials_Only", rngProf)

    'Copie les données vers un tableau - 2024-11-20 @ 14:19
    Dim rngData As Range
    Set rngData = ws.Range("A1").CurrentRegion
    Dim arr() As Variant
    Dim headerRows As Long: headerRows = 1
    Call TransfererPlageVersTableau2D(rngData, arr, 1)
    
    Dim dict_TECID As New Dictionary
    
    Dim i As Long, tecID As Long
    Dim clientCode As String
    Dim dateTEC As Date, minDate As Date, maxDate As Date
    Dim hres As Double, hres_non_detruites As Double
    Dim estDetruit As Boolean, estFacturable As Boolean, estFacturee As Boolean
    Dim cas_doublon_TecID As Long, cas_date_invalide As Long
    Dim cas_hres_invalide As Long, cas_estFacturable_invalide As Long
    Dim cas_estDetruit_invalide As Long, cas_estFacturee_invalide As Long
    Dim total_hres_inscrites As Double, total_hres_detruites As Double, total_hres_facturees As Double
    Dim total_hres_facturable As Double, total_hres_non_facturable As Double
    
    minDate = "12/31/2999"
    
    Dim isTECTDBValid As Boolean
    isTECTDBValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        tecID = arr(i, 1)
        If Fn_Is_String_Valid(CStr(arr(i, 3)), rngProf) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le professionnel '" & arr(i, 3) & "' est INVALIDE")
            r = r + 1
            isTECTDBValid = False
        End If
        dateTEC = arr(i, 4)
        If IsDate(dateTEC) = False Or arr(i, 4) > Date Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** TECID =" & tecID & " a une date INVALIDE '" & dateTEC & " !!!")
            r = r + 1
            isTECTDBValid = False
            cas_date_invalide = cas_date_invalide + 1
        Else
            If dateTEC < minDate Then minDate = dateTEC
            If dateTEC > maxDate Then maxDate = dateTEC
        End If
        If dateTEC <> Int(dateTEC) Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** La date du TEC '" & dateTEC & "' n'est pas du bon format (H:M:S) pour le TECID =" & tecID)
            r = r + 1
            isTECTDBValid = False
        End If
        clientCode = arr(i, 5)
        hres = arr(i, 8)
        If IsNumeric(hres) = False Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            r = r + 1
            isTECTDBValid = False
            cas_hres_invalide = cas_hres_invalide + 1
        End If
        estFacturable = arr(i, 9)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            r = r + 1
            isTECTDBValid = False
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If
        estFacturee = arr(i, 10)
        If InStr("Vrai^Faux^", estFacturee & "^") = 0 Or Len(estFacturee) <> 2 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            r = r + 1
            isTECTDBValid = False
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        estDetruit = arr(i, 11)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            r = r + 1
            isTECTDBValid = False
            cas_estDetruit_invalide = cas_estDetruit_invalide + 1
        End If
        
        'Heures Inscrites
        total_hres_inscrites = total_hres_inscrites + hres
        hres_non_detruites = hres
        
        'Heures détruites
        If estDetruit = "Vrai" Then
            total_hres_detruites = total_hres_detruites + hres
            hres_non_detruites = hres_non_detruites - hres
        End If
        
        'Heures FACTURABLES
        If hres_non_detruites <> 0 And estFacturable = "Vrai" And _
            Fn_Is_Client_Facturable(clientCode) = True Then
                total_hres_facturable = total_hres_facturable + hres_non_detruites
        End If
        
        'Heures non-FACTURABLES
        If hres_non_detruites <> 0 Then
            If estFacturable = "Faux" Or Fn_Is_Client_Facturable(clientCode) = False Then
                total_hres_non_facturable = total_hres_non_facturable + hres_non_detruites
            End If
        End If
        
        'Heures FACTURÉES
        If hres_non_detruites <> 0 And estDetruit = "Faux" And estFacturee = "Vrai" And _
            Fn_Is_Client_Facturable(clientCode) = True Then
                total_hres_facturees = total_hres_facturees + hres_non_detruites
        End If
        
        'Dictionary
        If dict_TECID.Exists(tecID) = False Then
            dict_TECID.Add tecID, 0
        Else
            Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Le TECID '" & tecID & "' est un doublon pour la ligne '" & i & "'")
            r = r + 1
            isTECTDBValid = False
           cas_doublon_TecID = cas_doublon_TecID + 1
        End If
        
    Next i
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " charges de temps ont été analysées!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    r = r + 1
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    r = r + 1
    
    If cas_doublon_TecID = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucun doublon de TECID")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_doublon_TecID & " cas de doublons pour les TECID")
        r = r + 1
        isTECTDBValid = False
    End If
    
    If cas_date_invalide = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucune date INVALIDE")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        r = r + 1
        isTECTDBValid = False
    End If
    
    If cas_hres_invalide = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucune heures INVALIDE")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        r = r + 1
        isTECTDBValid = False
    End If
    
    If cas_estFacturable_invalide = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucune valeur 'estFacturable' n'est INVALIDE")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        r = r + 1
        isTECTDBValid = False
    End If
    
    If cas_estFacturee_invalide = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucune valeur 'estFacturee' n'est INVALIDE")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        r = r + 1
        isTECTDBValid = False
    End If
    
    If cas_estDetruit_invalide = 0 Then
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Aucune valeur 'estDetruit' n'est INVALIDE")
        r = r + 1
    Else
        Call AjouterMessageAuxResultats(wsOutput, r, 2, "********** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        r = r + 1
        isTECTDBValid = False
    End If
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "La somme des heures saisies donne ces résultats:")
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Heures SAISIES         : " & formattedHours)
    gValeursAComparer(6, 1) = "Total des heures Saisies"
    gValeursAComparer(6, 2) = CCur(formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Heures détruites       : " & formattedHours)
    gValeursAComparer(7, 1) = "Total des heures détruites"
    gValeursAComparer(7, 2) = CCur(formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Heures NETTES          : " & formattedHours)
    gValeursAComparer(8, 1) = "Total des heures NETTES"
    gValeursAComparer(8, 2) = CCur(formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "              Non_facturables : " & formattedHours)
    gValeursAComparer(9, 1) = "Total des heures Non-Facturables"
    gValeursAComparer(9, 2) = CCur(formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "              Facturables     : " & formattedHours)
    gValeursAComparer(10, 1) = "Total des heures Facturables"
    gValeursAComparer(10, 2) = CCur(formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessageAuxResultats(wsOutput, r, 2, "       Heures facturées       : " & formattedHours)
    gValeursAComparer(11, 1) = "Total des heures facturées"
    gValeursAComparer(11, 2) = CCur(formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable - total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Heures TEC             : " & formattedHours
    rng.Characters(InStr(rng.Value, ":") + 2, Len(formattedHours)).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, ":") + 2, Len(formattedHours)).Font.Bold = True
    gValeursAComparer(12, 1) = "Total des heures TEC"
    gValeursAComparer(12, 2) = CCur(formattedHours)
    r = r + 2

    'Cas problème dans cette vérification ?
    If isTECTDBValid = False Then
        gverificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    On Error Resume Next
    Set rng = Nothing
    Set rngData = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierTECTdBData", vbNullString, startTime)

End Sub

Private Sub VerifierTEC(ByVal wsOutput As Worksheet, ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:VerifierTEC", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wsdTEC_Local
    
    Dim lastTECIDReported As Long
    lastTECIDReported = 7701 'What is the last TECID analyzed ?
    
    'Réference au UserDefined structure 'StatistiquesTEC'
    Dim stats As StatistiquesTEC
    
    'Feuille contenant les données à analyser
    Dim lo As ListObject
    Set lo = ws.ListObjects("l_tbl_TEC_Local")
    
    'Vérifier que le tableau existe bien
    If lo Is Nothing Then
        Call AjouterMessage(wsOutput, r, 2, "Tableau structuré 'tblTEC_Local' introuvable.")
        Exit Sub
    End If
    
    'Vérifier que le tableau contient au moins une ligne de données
    If lo.DataBodyRange Is Nothing Then
        Call AjouterMessage(wsOutput, r, 2, "Aucune donnée dans le tableau 'tblTEC_Local'.")
        Exit Sub
    End If
    
    Call AjouterMessage(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wsdTEC_Local'")
    Call AjouterMessage(wsOutput, r, 2, "Il y a " & Format$(lo.ListRows.count, "###,##0") & _
        " lignes et " & Format$(lo.ListColumns.count, "#,##0") & _
        " colonnes dans cette table")
    
    'Charger les données dans le tableau 2D
    Dim arrTEC_Local_Data As Variant
    arrTEC_Local_Data = lo.DataBodyRange.Value
    
    'Création de quelques dictionnaires pour accumuler les données
    Dim arr As Variant
    
    '1. Remplir le dictionnaire dictClient
    Set lo = wsdBD_Clients.ListObjects("l_tbl_BD_Clients")
    arr = lo.DataBodyRange.Value
    
    Dim dictClient As New Dictionary 'Dictionnaire pour tous les clients (code & nom)
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        If Not dictClient.Exists(arr(i, 1)) Then
            dictClient.Add arr(i, 2), arr(i, 1)
        End If
    Next i
    
    '2. Remplir le dictionnaire (dictFacture) pour toutes les factures émises
    Set lo = wsdFAC_Entete.ListObjects("l_tbl_FAC_Entête")
    arr = lo.DataBodyRange.Value
    
    Dim dictFacture As New Dictionary 'Dictionnaire pour toutes les factures émises
    For i = 1 To UBound(arr, 1)
        If Not dictFacture.Exists(arr(i, fFacCCInvNo)) Then
            dictFacture.Add arr(i, fFacCCInvNo), arr(i, fFacCCInvoiceDate)
        End If
    Next i
    
    'Créer des dictionnaires vides pour accumuler des informations
    Dim dictDateCharge As New Dictionary
    Dim dictDict As New Dictionary
    Dim dictFactureHres As New Dictionary
    Dim dictProf As New Dictionary
    Dim dictTimeStamp As New Dictionary
    
    Dim minDate As Date, maxDate As Date
    Dim maxTECID As Long
    Dim isTECvalid As Boolean

    minDate = "12/31/2999"
    
    'Boucle principale - Lecture et analyse des TEC (TEC_Local)
    For i = LBound(arrTEC_Local_Data, 1) To UBound(arrTEC_Local_Data, 1)
    
        isTECvalid = True
        If Not AnalyserLigneTEC(arrTEC_Local_Data, i, wsOutput, r, _
                                minDate, maxDate, _
                                dictClient, dictFacture, dictFactureHres, _
                                dictDateCharge, dictTimeStamp, dictDict, dictProf, _
                                stats, lastTECIDReported) Then
            isTECvalid = False
            stats.nbInvalid = stats.nbInvalid + 1
        Else
            stats.nbValid = stats.nbValid + 1
        End If
        
        'Détermine le plus grand TecID
        If arrTEC_Local_Data(i, fTECTECID) > maxTECID Then
            maxTECID = arrTEC_Local_Data(i, fTECTECID)
        End If

    Next i
        
    'Add number of rows processed (read)
    readRows = readRows + UBound(arrTEC_Local_Data, 1)
    
    'Imprimer la conclusion des tests
    Call ImprimerCasProbleme(arrTEC_Local_Data, wsOutput, r, minDate, maxDate, _
                             stats, isTECvalid)
                                
    Call ComparerHeuresFactureesSelon2Sources(wsOutput, r, dictFacture, dictFactureHres, isTECvalid)
    
    Call ImprimerSommaireHeuresTEC(wsOutput, r, stats)
    
    Call ImprimerSommaireDateProf(wsOutput, r, dictDateCharge, maxTECID)
    
    Call ImprimerSommaireTimeStampProf(wsOutput, r, dictTimeStamp, lastTECIDReported, isTECvalid)
    
    'Cas problème dans cette vérification ?
    If isTECvalid = False Then
        gverificationIntegriteOK = False
    End If
    
    Call AjouterMessage(wsOutput, r, 2, "Temps d’exécution : " & Format(Timer - startTime, "0.00") & " seconde(s)")
    r = r + 1

    'Libérer la mémoire (test)
    On Error Resume Next
    Set dictClient = Nothing
    Set dictDateCharge = Nothing
    Set dictFacture = Nothing
    Set dictTimeStamp = Nothing
    Set dictProf = Nothing
    Set dictDict = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    Call EnregistrerLogApplication("modAppli_Utils:VerifierTEC", vbNullString, startTime)

End Sub

Sub ImprimerCasProbleme(data As Variant, ByVal wsOutput As Worksheet, ByRef r As Long, _
                        ByRef minDate As Date, ByRef maxDate As Date, _
                        ByRef stats As StatistiquesTEC, isTECvalid As Boolean)

    Call AjouterMessage(wsOutput, r, 2, "Un total de " & Format$(UBound(data, 1), "##,##0") & " charges de temps ont été analysées!")
    
    'Impression du sommaire de l'analyse
    Call AjouterMessage(wsOutput, r, 2, "       La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    Call AjouterMessage(wsOutput, r, 2, "       La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    
    If stats.cas_doublon_TecID = 0 Then
        Call AjouterMessage(wsOutput, r, 2, "       Aucun doublon de TECID")
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Il y a " & stats.cas_doublon_TecID & " cas de doublons pour les TecID")
        isTECvalid = False
    End If
    
    If stats.cas_date_invalide = 0 Then
        Call AjouterMessage(wsOutput, r, 2, "       Aucune date INVALIDE")
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Il y a " & stats.cas_date_invalide & " cas de date INVALIDE")
        isTECvalid = False
    End If
    
    If stats.cas_date_future = 0 Then
        Call AjouterMessage(wsOutput, r, 2, "       Aucune date dans le futur")
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Il y a " & stats.cas_date_future & " cas de date FUTURE")
        isTECvalid = False
    End If
    
    If stats.cas_hres_invalide = 0 Then
        Call AjouterMessage(wsOutput, r, 2, "       Aucune heures INVALIDE")
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Il y a " & stats.cas_hres_invalide & " cas d'heures INVALIDE")
        isTECvalid = False
    End If
    
    If stats.cas_estFacturable_invalide = 0 Then
        Call AjouterMessage(wsOutput, r, 2, "       Aucune valeur 'estFacturable' n'est INVALIDE")
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Il y a " & stats.cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        isTECvalid = False
    End If
    
    If stats.cas_estFacturee_invalide = 0 Then
        Call AjouterMessage(wsOutput, r, 2, "       Aucune valeur 'estFacturee' n'est INVALIDE")
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Il y a " & stats.cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        isTECvalid = False
    End If
    
    If stats.cas_estDetruit_invalide = 0 Then
        Call AjouterMessage(wsOutput, r, 2, "       Aucune valeur 'estDetruit' n'est INVALIDE")
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Il y a " & stats.cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        isTECvalid = False
    End If
    
    If stats.cas_date_fact_invalide = 0 Then
        Call AjouterMessage(wsOutput, r, 2, "       Aucune date de facture INVALIDE")
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Il y a " & stats.cas_date_fact_invalide & " cas de date de facture INVALIDE")
        isTECvalid = False
    End If

End Sub

Sub ComparerHeuresFactureesSelon2Sources(ByVal wsOutput As Worksheet, ByRef r As Long, _
                                         ByRef dictFacture As Object, ByRef dictFactureHres As Object, _
                                         ByRef isTECvalid As Boolean)

    Call AjouterMessage(wsOutput, r, 2, "Vérification des Heures Facturées par Facture")

    'Vérification des heures facturées selon 2 sources (TEC_Local vs. FAC_Détails)
    Dim key As Variant
    Dim totalHoursBilled As Currency
    Dim hresFactureesTEC_Local As Currency
    Dim cas_Heures_Differentes As Integer
    
    For Each key In dictFacture.keys
        totalHoursBilled = Fn_Get_TEC_Total_Invoice_AF(CStr(key), "Heures")
        hresFactureesTEC_Local = dictFactureHres(key)
        If Not EgalCurrency(totalHoursBilled, hresFactureesTEC_Local) Then
            Call AjouterMessage(wsOutput, r, 2, "********** Facture '" & CStr(key) & _
                    "', il y a un écart d'heures facturées entre TEC_Local & FAC_Détails - " & _
                        Round(hresFactureesTEC_Local, 2) & " vs. " & Round(totalHoursBilled, 2))
            isTECvalid = False
            cas_Heures_Differentes = cas_Heures_Differentes + 1
        End If
    Next key

    If cas_Heures_Differentes = 0 Then
        Call AjouterMessage(wsOutput, r, 2, "       Toutes les heures facturées balancent, selon les 2 sources")
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Certaines factures sont à vérifier pour que les heures facturées balancent, selon les 2 sources")
        isTECvalid = False
    End If
        
End Sub

Sub ImprimerSommaireHeuresTEC(ByVal wsOutput As Worksheet, ByRef r As Long, _
                              ByRef stats As StatistiquesTEC)
    
    Call AjouterMessage(wsOutput, r, 2, "La somme des heures SAISIES donne ces résultats:")
    
    Dim formattedHours As String
    formattedHours = Format$(stats.total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessage(wsOutput, r, 2, "       Heures SAISIES         : " & formattedHours)
    gValeursAComparer(6, 3) = CCur(formattedHours)
    
    formattedHours = Format$(stats.total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessage(wsOutput, r, 2, "       Heures détruites       : " & formattedHours)
    gValeursAComparer(7, 3) = CCur(formattedHours)
    
    formattedHours = Format$(stats.total_hres_inscrites - stats.total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessage(wsOutput, r, 2, "       Heures NETTES          : " & formattedHours)
    gValeursAComparer(8, 3) = CCur(formattedHours)
    
    formattedHours = Format$(stats.total_hres_non_facturables, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessage(wsOutput, r, 2, "              Non_facturables : " & formattedHours)
    gValeursAComparer(9, 3) = CCur(formattedHours)

    formattedHours = Format$(stats.total_hres_facturables, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessage(wsOutput, r, 2, "              Facturables     : " & formattedHours)
    gValeursAComparer(10, 3) = CCur(formattedHours)
    
    formattedHours = Format$(stats.total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AjouterMessage(wsOutput, r, 2, "       Heures facturées       : " & formattedHours)
    gValeursAComparer(11, 3) = CCur(formattedHours)

    formattedHours = Format$(stats.total_hres_facturables - stats.total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Heures TEC             : " & formattedHours
    rng.Characters(InStr(rng.Value, ":") + 2, Len(formattedHours)).Font.Color = vbRed
    rng.Characters(InStr(rng.Value, ":") + 2, Len(formattedHours)).Font.Bold = True
    gValeursAComparer(12, 3) = CCur(formattedHours)
    r = r + 2
    
End Sub

Sub ImprimerSommaireDateProf(ByVal wsOutput As Worksheet, ByRef r As Long, _
                             ByRef dictDateCharge As Object, maxTECID As Long)
    
    Dim formattedHours As String
    Dim key As Variant
    Dim keys() As Variant
    Dim rng As Range
    
    'Tri & impression de dictDateCharge
    If dictDateCharge.count > 0 Then
        'Avec un peu de couleur
        Set rng = wsOutput.Range("B" & r)
        rng.Value = "Sommaire des heures selon la DATE de la charge (" & maxTECID & ")"
        rng.Characters(InStr(rng.Value, "(") + 1, Len(maxTECID)).Font.Color = vbGreen
        rng.Characters(InStr(rng.Value, "(") + 1, Len(maxTECID)).Font.Bold = True
        r = r + 1
        
        keys = dictDateCharge.keys
        Call Fn_Quick_Sort(keys, LBound(keys), UBound(keys))
        'Parcourir les clés triées et afficher les heures
        Dim i As Long
        For i = LBound(keys) To UBound(keys)
            key = keys(i)
            formattedHours = Format$(dictDateCharge(key), "#0.00")
            formattedHours = String(6 - Len(formattedHours), " ") & formattedHours
            Call AjouterMessage(wsOutput, r, 2, "       " & key & ":" & formattedHours & " heures")
        Next i
    End If

End Sub

Sub ImprimerSommaireTimeStampProf(ByVal wsOutput As Worksheet, ByRef r As Long, _
                                  ByRef dictTimeStamp As Object, lastTECIDReported As Long, _
                                  ByRef isTECvalid As Boolean)
    
    Dim keys As Variant
    Dim key As Variant
    Dim formattedHours As String
    'Tri & impression de dictTimeStamp
    If dictTimeStamp.count > 0 Then
        Call AjouterMessage(wsOutput, r, 2, "Sommaire des heures saisies selon le 'TIMESTAMP'")
        keys = dictTimeStamp.keys
        Call Fn_Quick_Sort(keys, LBound(keys), UBound(keys))
        'Parcourir les clés triées et afficher les valeurs
        Dim i As Long
        For i = LBound(keys) To UBound(keys)
            key = keys(i)
            formattedHours = Format$(dictTimeStamp(key), "##0")
            formattedHours = String(6 - Len(formattedHours), " ") & formattedHours
            Call AjouterMessage(wsOutput, r, 2, "       " & key & ":" & formattedHours & " entrée(s)")
        Next i
    Else
        Call AjouterMessage(wsOutput, r, 2, "Aucune nouvelle saisie d'heures (TECID > " & lastTECIDReported & ") ")
        r = r + 1
    End If

End Sub

Sub Make_It_As_Header(r As Range, couleurFond As Long) '2025-06-30 @ 14:08

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
    MsgBox "Erreur dans Make_It_As_Header : " & Err.description, vbExclamation, "Erreur VBA"
    
End Sub

Sub AjouterMessageAuxResultats(ws As Worksheet, r As Long, c As Long, m As String)

    ws.Cells(r, c).Value = m
    If c = 1 Then
        ws.Cells(r, c).Font.Bold = True
    End If

End Sub

Sub AppliquerConditionalFormating(rng As Range, headerRows As Long, couleurFond As Long)

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:AppliquerConditionalFormating", vbNullString, 0)
    
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
    
    Call EnregistrerLogApplication("modAppli_Utils:AppliquerConditionalFormating", vbNullString, startTime)

End Sub

Sub AppliquerFormatColonnesParTable(ws As Worksheet, rng As Range, HeaderRow As Long)

    'Conditional Formatting (many steps)
    '1) Remove existing conditional formatting
'        rng.Cells.FormatConditions.Delete 'Remove the worksheet conditional formatting
    
    '2) Define the usedRange to data only (exclude header row(s))
        Dim numRows As Long
        numRows = rng.CurrentRegion.Rows.count - HeaderRow
        Dim usedRange As Range
        If numRows > 0 Then
            On Error Resume Next
            Set usedRange = rng.offset(HeaderRow, 0).Resize(numRows, rng.Columns.count)
            On Error GoTo 0
        End If
        
    '3) Specific columns formats to worksheets
        Dim lastUsedRow As Long
        lastUsedRow = rng.Rows.count
        If lastUsedRow = HeaderRow Then
            Exit Sub
        End If
        
        Dim rngUnion As Range
        
        Select Case rng.Worksheet.CodeName
            Case "wsdCC_Regularisations" '2025-02-12 @ 07:58
                With wsdCC_Regularisations
                    .Range(.Cells(2, fREGULRegulID), .Cells(lastUsedRow, fREGULTimeStamp)).HorizontalAlignment = xlCenter
                    .Range(.Cells(2, fREGULInvNo), .Cells(lastUsedRow, fREGULInvNo)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(2, fREGULClientNom), .Cells(lastUsedRow, fREGULClientNom)).HorizontalAlignment = xlLeft
                    .Range(.Cells(2, fREGULDescription), .Cells(lastUsedRow, fREGULDescription)).HorizontalAlignment = xlLeft
                    With .Range(.Cells(2, fREGULHono), .Cells(lastUsedRow, fREGULTVQ))
                        .HorizontalAlignment = xlRight
                        .NumberFormat = "#,##0.00"
                    End With
                    .Range(.Cells(2, fREGULTimeStamp), .Cells(lastUsedRow, fREGULTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
                    
            Case "wsdDEB_Recurrent"  '2025-02-12 @ 08:06
                With wsdDEB_Recurrent
                    .Range(.Cells(2, fDebRNoDebRec), .Cells(lastUsedRow, fDebRTimeStamp)).HorizontalAlignment = xlCenter
                    .Range(.Cells(2, fDebRDate), .Cells(lastUsedRow, fDebRDate)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(2, fDebRType), .Cells(lastUsedRow, fDebRReference)).HorizontalAlignment = xlLeft
                    .Range(.Cells(2, fDebRCompte), .Cells(lastUsedRow, fDebRCompte)).HorizontalAlignment = xlLeft
                    With .Range(.Cells(2, fDebRTotal), .Cells(lastUsedRow, fDebRCréditTVQ))
                        .HorizontalAlignment = xlRight
                        .NumberFormat = "#,##0.00"
                    End With
                    .Range(.Cells(2, fDebRTimeStamp), .Cells(lastUsedRow, fDebRTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
           
            Case "wsdDEB_Trans" '2025-02-12 @ 08:22
                With wsdDEB_Trans
                    Set rngUnion = Application.Union( _
                        .Range(ws.Cells(2, fDebTType), ws.Cells(lastUsedRow, fDebTBeneficiaire)), _
                        .Range(ws.Cells(2, fDebTDescription), ws.Cells(lastUsedRow, fDebTReference)), _
                        .Range(ws.Cells(2, fDebTCompte), ws.Cells(lastUsedRow, fDebTCompte)), _
                        .Range(ws.Cells(2, fDebTAutreRemarque), ws.Cells(lastUsedRow, fDebTAutreRemarque)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlLeft
                    .Range(.Cells(2, fDebTNoEntrée), .Cells(lastUsedRow, fDebTDate)).HorizontalAlignment = xlCenter
                    .Range(.Cells(2, fDebTDate), .Cells(lastUsedRow, fDebTDate)).NumberFormat = "yyyy-mm-dd"
                    With .Range(.Cells(2, fDebTTotal), .Cells(lastUsedRow, fDebTDépense))
                        .HorizontalAlignment = xlRight
                        .NumberFormat = "#,##0.00"
                    End With
                    .Range(.Cells(2, fDebTTimeStamp), .Cells(lastUsedRow, fDebTTimeStamp)).HorizontalAlignment = xlCenter
                    .Range(.Cells(2, fDebTTimeStamp), .Cells(lastUsedRow, fDebTTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
            
            Case "wsdENC_Details" '2025-02-12 @ 08:33
                With wsdENC_Details
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(2, fEncDPayID), .Cells(lastUsedRow, fEncDPayID)), _
                        .Range(.Cells(2, fEncDInvNo), .Cells(lastUsedRow, fEncDInvNo)), _
                        .Range(.Cells(2, fEncDPayDate), .Cells(lastUsedRow, fEncDPayDate)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlCenter
                    .Range(.Cells(2, fEncDCustomer), .Cells(lastUsedRow, fEncDCustomer)).HorizontalAlignment = xlLeft
                    With .Range(.Cells(2, fEncDPayAmount), .Cells(lastUsedRow, fEncDPayAmount))
                        .HorizontalAlignment = xlRight
                        .NumberFormat = "#,##0.00"
                    End With
                    .Range(.Cells(2, fEncDPayID), .Cells(lastUsedRow, fEncDPayID)).NumberFormat = "0"
                    .Range(.Cells(2, fEncDPayDate), .Cells(lastUsedRow, fEncDPayDate)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(2, fEncDTimeStamp), .Cells(lastUsedRow, fEncDTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
            
            Case "wsdENC_Entete" '2025-02-12 @ 08:39
                With wsdENC_Entete
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(2, fEncEPayID), .Cells(lastUsedRow, fEncEPayID)), _
                        .Range(.Cells(2, fEncEPayDate), .Cells(lastUsedRow, fEncEPayDate)), _
                        .Range(.Cells(2, fEncECodeClient), .Cells(lastUsedRow, fEncECodeClient)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlCenter
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(2, fEncECustomer), .Cells(lastUsedRow, fEncECustomer)), _
                        .Range(.Cells(2, fEncEPayType), .Cells(lastUsedRow, fEncEPayType)), _
                        .Range(.Cells(2, fEncENotes), .Cells(lastUsedRow, fEncENotes)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlLeft
                    .Range(.Cells(2, fEncEAmount), .Cells(lastUsedRow, fEncEAmount)).HorizontalAlignment = xlRight
                    .Range(.Cells(2, fEncEPayID), .Cells(lastUsedRow, fEncEPayID)).NumberFormat = "0"
                    .Range(.Cells(2, fEncEPayDate), .Cells(lastUsedRow, fEncEPayDate)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(2, fEncEAmount), .Cells(lastUsedRow, fEncEAmount)).NumberFormat = "#,##0.00"
                    .Range(.Cells(2, fEncETimeStamp), .Cells(lastUsedRow, fEncETimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
            
            Case "wsdFAC_Comptes_Clients" '2025-01-25 @ 15:35
                With wsdFAC_Comptes_Clients
                    .Range(.Cells(3, fFacCCInvNo), .Cells(lastUsedRow, fFacCCInvoiceDate)).HorizontalAlignment = xlCenter
                    .Range(.Cells(3, fFacCCCodeClient), .Cells(lastUsedRow, fFacCCDueDate)).HorizontalAlignment = xlCenter
                    .Range(.Cells(3, fFacCCCustomer), .Cells(lastUsedRow, fFacCCCustomer)).HorizontalAlignment = xlLeft
                    .Range(.Cells(3, fFacCCTotal), .Cells(lastUsedRow, fFacCCBalance)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacCCInvoiceDate), .Cells(lastUsedRow, fFacCCInvoiceDate)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(3, fFacCCDueDate), .Cells(lastUsedRow, fFacCCDueDate)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(3, fFacCCTotal), .Cells(lastUsedRow, fFacCCBalance)).NumberFormat = "###,##0.00"
                    .Range(.Cells(3, fFacCCTimeStamp), .Cells(lastUsedRow, fFacCCTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
            
            Case "wsdFAC_Details" '2025-02-12 @ 10:15
                With wsdFAC_Details
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(3, fFacDInvNo), .Cells(lastUsedRow, fFacDInvNo)), _
                        .Range(.Cells(3, fFacDInvRow), .Cells(lastUsedRow, fFacDInvRow)), _
                        .Range(.Cells(3, fFacDTimeStamp), .Cells(lastUsedRow, fFacDTimeStamp)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlCenter
                    .Range(.Cells(3, fFacDDescription), .Cells(lastUsedRow, fFacDDescription)).HorizontalAlignment = xlLeft
                    .Range(.Cells(3, fFacDHeures), .Cells(lastUsedRow, fFacDHonoraires)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacDHeures), .Cells(lastUsedRow, fFacDHonoraires)).NumberFormat = "#,##0.00"
                    .Range(.Cells(3, fFacDTimeStamp), .Cells(lastUsedRow, fFacDTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
            
            Case "wsdFAC_Entete" '2025-02-12 @ 10:36
                With wsdFAC_Entete
                    .Range(.Cells(3, fFacEInvNo), .Cells(lastUsedRow, fFacECustID)).HorizontalAlignment = xlCenter
                    .Range(.Cells(3, fFacETimeStamp), .Cells(lastUsedRow, fFacETimeStamp)).HorizontalAlignment = xlCenter
                    .Range(.Cells(3, fFacEContact), .Cells(lastUsedRow, fFacEAdresse3)).HorizontalAlignment = xlLeft
                    .Range(.Cells(3, fFacEAF1Desc), .Cells(lastUsedRow, fFacEAF1Desc)).HorizontalAlignment = xlLeft
                    .Range(.Cells(3, fFacEAF2Desc), .Cells(lastUsedRow, fFacEAF2Desc)).HorizontalAlignment = xlLeft
                    .Range(.Cells(3, fFacEAF3Desc), .Cells(lastUsedRow, fFacEAF3Desc)).HorizontalAlignment = xlLeft
                    .Range(.Cells(3, fFacEHonoraires), .Cells(lastUsedRow, fFacEHonoraires)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacEHonoraires), .Cells(lastUsedRow, fFacEHonoraires)).NumberFormat = "#,##0.00"
                    .Range(.Cells(3, fFacEAutresFrais1), .Cells(lastUsedRow, fFacEAutresFrais1)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacEAutresFrais1), .Cells(lastUsedRow, fFacEAutresFrais1)).NumberFormat = "#,##0.00"
                    .Range(.Cells(3, fFacEAutresFrais2), .Cells(lastUsedRow, fFacEAutresFrais2)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacEAutresFrais2), .Cells(lastUsedRow, fFacEAutresFrais2)).NumberFormat = "#,##0.00"
                    .Range(.Cells(3, fFacEAutresFrais3), .Cells(lastUsedRow, fFacEAutresFrais3)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacEAutresFrais3), .Cells(lastUsedRow, fFacEAutresFrais3)).NumberFormat = "#,##0.00"
                    .Range(.Cells(3, fFacEMntTPS), .Cells(lastUsedRow, fFacEMntTPS)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacEMntTPS), .Cells(lastUsedRow, fFacEMntTPS)).NumberFormat = "#,##0.00"
                    .Range(.Cells(3, fFacEMntTVQ), .Cells(lastUsedRow, fFacEMntTVQ)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacEMntTVQ), .Cells(lastUsedRow, fFacEMntTVQ)).NumberFormat = "#,##0.00"
                    .Range(.Cells(3, fFacEARTotal), .Cells(lastUsedRow, fFacEARTotal)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacEARTotal), .Cells(lastUsedRow, fFacEARTotal)).NumberFormat = "#,##0.00"
                    .Range(.Cells(3, fFacEDépôt), .Cells(lastUsedRow, fFacEDépôt)).HorizontalAlignment = xlRight
                    .Range(.Cells(3, fFacEDépôt), .Cells(lastUsedRow, fFacEDépôt)).NumberFormat = "#,##0.00"
                    .Range(.Cells(3, fFacEDateFacture), .Cells(lastUsedRow, fFacEDateFacture)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(3, fFacETauxTPS), .Cells(lastUsedRow, fFacETauxTPS)).HorizontalAlignment = xlCenter
                    .Range(.Cells(3, fFacETauxTPS), .Cells(lastUsedRow, fFacETauxTPS)).NumberFormat = "#0.000 %"
                    .Range(.Cells(3, fFacETauxTVQ), .Cells(lastUsedRow, fFacETauxTVQ)).HorizontalAlignment = xlCenter
                    .Range(.Cells(3, fFacETauxTVQ), .Cells(lastUsedRow, fFacETauxTVQ)).NumberFormat = "#0.000 %"
                    .Range(.Cells(3, fFacETimeStamp), .Cells(lastUsedRow, fFacETimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
    
            Case "wsdFAC_Projets_Details" '2025-02-12 @ 11:42
                With wsdFAC_Projets_Details
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(2, fFacPDProjetID), .Cells(lastUsedRow, fFacPDProjetID)) _
                        .Range(.Cells(2, fFacPDClientID), .Cells(lastUsedRow, fFacPDProf)), _
                        .Range(.Cells(2, fFacPDestDetruite), .Cells(lastUsedRow, fFacPDTimeStamp)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlCenter
                    .Range(.Cells(2, fFacPDNomClient), .Cells(lastUsedRow, fFacPDNomClient)).HorizontalAlignment = xlLeft
                    .Range(.Cells(2, fFacPDHeures), .Cells(lastUsedRow, fFacPDHeures)).HorizontalAlignment = xlRight
                    .Range(.Cells(2, fFacPDDate), .Cells(lastUsedRow, fFacPDDate)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(2, fFacPDHeures), .Cells(lastUsedRow, fFacPDHeures)).NumberFormat = "#,##0.00"
                    .Range(.Cells(2, fFacPDTimeStamp), .Cells(lastUsedRow, fFacPDTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
            
            Case "wsdFAC_Projets_Entete" '2025-02-12 @ 12:41
                With wsdFAC_Projets_Entete
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(2, fFacPEProjetID), .Cells(lastUsedRow, fFacPEProjetID)) _
                        .Range(.Cells(2, fFacPEClientID), .Cells(lastUsedRow, fFacPEDate)), _
                        .Range(.Cells(2, fFacPEProf1), .Cells(lastUsedRow, fFacPEProf1)), _
                        .Range(.Cells(2, fFacPEProf2), .Cells(lastUsedRow, fFacPEProf2)), _
                        .Range(.Cells(2, fFacPEProf3), .Cells(lastUsedRow, fFacPEProf3)), _
                        .Range(.Cells(2, fFacPEProf4), .Cells(lastUsedRow, fFacPEProf4)), _
                        .Range(.Cells(2, fFacPEProf5), .Cells(lastUsedRow, fFacPEProf5)), _
                        .Range(.Cells(2, fFacPEestDetruite), .Cells(lastUsedRow, fFacPETimeStamp)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlCenter
                    .Range(.Cells(2, fFacPENomClient), .Cells(lastUsedRow, fFacPENomClient)).HorizontalAlignment = xlLeft
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(2, fFacPEHonoTotal), .Cells(lastUsedRow, fFacPEHonoTotal)) _
                        .Range(.Cells(2, fFacPEHres1), .Cells(lastUsedRow, fFacPEHono1)), _
                        .Range(.Cells(2, fFacPEHres2), .Cells(lastUsedRow, fFacPEHono2)), _
                        .Range(.Cells(2, fFacPEHres3), .Cells(lastUsedRow, fFacPEHono3)), _
                        .Range(.Cells(2, fFacPEHres4), .Cells(lastUsedRow, fFacPEHono4)), _
                        .Range(.Cells(2, fFacPEHres5), .Cells(lastUsedRow, fFacPEHono5)))
                     If Not rngUnion Is Nothing Then
                        rngUnion.HorizontalAlignment = xlRight
                        rngUnion.NumberFormat = "###,##0.00"
                    End If
                End With
            
            Case "wsdFAC_Sommaire_Taux" '2025-02-12 @ 12:50
                With wsdFAC_Sommaire_Taux
                    .Range(.Cells(2, fFacSTInvNo), .Cells(lastUsedRow, fFacSTProf)).HorizontalAlignment = xlCenter
                    .Range(.Cells(2, fFacSTHeures), .Cells(lastUsedRow, fFacSTTaux)).HorizontalAlignment = xlRight
                    .Range(.Cells(2, fFacSTHeures), .Cells(lastUsedRow, fFacSTTaux)).NumberFormat = "#,##0.00"
                    .Range(.Cells(2, fFacSTTimeStamp), .Cells(lastUsedRow, fFacSTTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
            
            Case "wsdGL_EJ_Recurrente" '2025-02-12 @ 12:59
                With wsdGL_EJ_Recurrente
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(2, fGlEjRNoEjR), .Cells(lastUsedRow, fGlEjRNoEjR)) _
                        .Range(.Cells(2, fGlEjRNoCompte), .Cells(lastUsedRow, fGlEjRNoCompte)), _
                        .Range(.Cells(2, fGlEjRTimeStamp), .Cells(lastUsedRow, fGlEjRTimeStamp)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlCenter
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(2, fGlEjRDescription), .Cells(lastUsedRow, fGlEjRDescription)) _
                        .Range(.Cells(2, fGlEjRCompte), .Cells(lastUsedRow, fGlEjRCompte)), _
                        .Range(.Cells(2, fGlEjRAutreRemarque), .Cells(lastUsedRow, fGlEjRAutreRemarque)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlLeft
                    .Range(.Cells(2, fGlEjRDébit), .Cells(lastUsedRow, fGlEjRCrédit)).HorizontalAlignment = xlRight
                    .Range(.Cells(2, fGlEjRDébit), .Cells(lastUsedRow, fGlEjRCrédit)).NumberFormat = "###,##0.00 $"
                    .Range(.Cells(2, fGlEjRTimeStamp), .Cells(lastUsedRow, fGlEjRTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
            
            Case "wsdGL_Trans" '2025-02-12 @ 13:06
                With wsdGL_Trans
                    .Range(.Cells(2, fGlTNoCompte), .Cells(lastUsedRow, fGlTTimeStamp)).HorizontalAlignment = xlCenter
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(2, fGlTDate), .Cells(lastUsedRow, fGlTSource)), _
                        .Range(.Cells(2, fGlTCompte), .Cells(lastUsedRow, fGlTCompte)), _
                        .Range(.Cells(2, fGlTAutreRemarque), .Cells(lastUsedRow, fGlTAutreRemarque)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlLeft
                    .Range(.Cells(2, fGlTDébit), .Cells(lastUsedRow, fGlTCrédit)).HorizontalAlignment = xlRight
                    .Range(.Cells(2, fGlTDébit), .Cells(lastUsedRow, fGlTCrédit)).NumberFormat = "###,##0.00"
                    .Range(.Cells(2, fGlTDate), .Cells(lastUsedRow, fGlTDate)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(2, fGlTTimeStamp), .Cells(lastUsedRow, fGlTTimeStamp)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End With
            
            Case "wsdTEC_Local" '2025-02-12 @ 13:14
                With wsdTEC_Local
                    .Range(.Cells(3, fTECTECID), .Cells(lastUsedRow, fTECNoFacture)).HorizontalAlignment = xlCenter
                    Set rngUnion = Application.Union( _
                        .Range(.Cells(3, fTECClientNom), .Cells(lastUsedRow, fTECDescription)), _
                        .Range(.Cells(3, fTECCommentaireNote), .Cells(lastUsedRow, fTECCommentaireNote)), _
                        .Range(.Cells(3, fTECVersionApp), .Cells(lastUsedRow, fTECVersionApp)))
                    If Not rngUnion Is Nothing Then rngUnion.HorizontalAlignment = xlLeft
                    .Range(.Cells(3, fTECHeures), .Cells(lastUsedRow, fTECHeures)).NumberFormat = "#0.00"
                    .Range(.Cells(3, fTECDate), .Cells(lastUsedRow, fTECDate)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(3, fTECDateFacturee), .Cells(lastUsedRow, fTECDateFacturee)).NumberFormat = "yyyy-mm-dd"
                    .Range(.Cells(3, fTECDateSaisie), .Cells(lastUsedRow, fTECDateSaisie)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                    .Columns(fTECClientNom).ColumnWidth = 40
                    .Columns(fTECDescription).ColumnWidth = 55
                    .Columns(fTECCommentaireNote).ColumnWidth = 20
                End With
        End Select

    '4) Common stuff to all worksheets
        rng.EntireColumn.AutoFit
        rng.RowHeight = 15
        
        'Exceptions pour la largeur de colonne - 2025-07-09 @ 06:32
        If rng.Worksheet.CodeName = "wsdTEC_Local" Then
            With wsdTEC_Local
                .Columns(fTECClientNom).ColumnWidth = 45
                .Columns(fTECDescription).ColumnWidth = 65
                .Columns(fTECCommentaireNote).ColumnWidth = 30
            End With
        End If
    
    'Libérer la mémoire
    On Error Resume Next
    Set rngUnion = Nothing
    Set usedRange = Nothing
    On Error GoTo 0
    
End Sub

Sub Fix_Font_Size_And_Family(r As Range, ff As String, fs As Long)

    With r.Font
        .Name = ff
        .size = fs
        .underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With

End Sub

Sub Get_Deplacements_From_TEC()  '2024-09-05 @ 10:22

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:Get_Deplacements_From_TEC", vbNullString, 0)
    
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
    Call Make_It_As_Header(wsOutput.Range("A1:K1"), RGB(0, 112, 192))
    
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
    
    'Analyse de TEC_Local
    Call modImport.ImporterTEC
    
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
                clientData = Fn_Rechercher_Client_Par_ID(Trim$(arr(i, 5)), wsMF)
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
'    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:I" & rowOutput)
    Dim header1 As String: header1 = "Liste des TEC pour Guillaume"
    Dim header2 As String: header2 = "Période du " & dateFrom & " au " & dateTo
    Call MettreEnFormeImpressionSimple(wsOutput, rngArea, header1, header2, "$1:$1", "P")
    
    'Libérer la mémoire
    Set rngArea = Nothing
    Set rngClientsMF = Nothing
    Set wsOutput = Nothing
    Set wsMF = Nothing
    Set wsTEC = Nothing
    
    Call EnregistrerLogApplication("modAppli_Utils:Get_Deplacements_From_TEC", vbNullString, startTime)

End Sub

Sub Get_Date_Derniere_Modification(fileName As String, ByRef ddm As Date, _
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

Sub RedefinirDnrPlanComptable() '2024-07-04 @ 10:39
    
    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:RedefinirDnrPlanComptable", vbNullString, 0)

    'Redefine - dnrPlanComptable_Description_Only
    'Delete existing dynamic named range (assuming it could exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptable_Description_Only").Delete
    On Error GoTo 0
    
    'Define a new dynamic named range for 'dnrPlanComptable_Description_Only'
    Dim newRangeFormula As String
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,1)"
    
    'Create the new dynamic named range
    ThisWorkbook.Names.Add Name:="dnrPlanComptable_Description_Only", RefersTo:=newRangeFormula
    
    'Redefine - dnrPlanComptable_All
    'Delete existing dynamic named range (assuming it could exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptable_All").Delete
    On Error GoTo 0
    
    'Define a new dynamic named range for 'dnrPlanComptable_All'
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,4)"
    
    'Create the new dynamic named range
    ThisWorkbook.Names.Add Name:="dnrPlanComptable_All", RefersTo:=newRangeFormula
    
    Call EnregistrerLogApplication("modAppli_Utils:RedefinirDnrPlanComptable", vbNullString, startTime)

End Sub

Sub Remplir_Plage_Avec_Couleur(ByVal plage As Range, ByVal couleurRVB As Long)

    If Not plage Is Nothing Then
        Dim cellule As Range
        'Parcourt toutes les cellules de la plage (contiguës ou non)
        For Each cellule In plage
            On Error Resume Next
            cellule.Interior.Color = couleurRVB
            On Error GoTo 0
        Next cellule
    Else
        MsgBox "La plage spécifiée est invalide.", vbExclamation, "Procédure 'Remplir_Plage_Avec_Couleur'"
    End If
    
End Sub

Sub Paint_A_Range(rng As Range, colorRGB As String)
   
    Dim cell As Variant
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With

End Sub

Sub NoterNombreLignesParFeuille() '2025-01-22 @ 16:19

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modAppli_Utils:NoterNombreLignesParFeuille", vbNullString, 0)
    
    'Spécifiez les chemins des classeurs
    Dim cheminClasseurUsage As String
    cheminClasseurUsage = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & "GCF_File_Usage.xlsx"
    Dim cheminClasseurMASTER As String
    cheminClasseurMASTER = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & "GCF_BD_MASTER.xlsx"
    
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
    wsUsage.Cells(LigneDisponible, 1).Value = Format$(dateHeure, "yyyy-mm-dd hh:mm:ss")
    
    'Parcourir les cols de la première ligne pour les noms de feuilles
    Dim feuilleNom As String
    Dim lastUsedRow As Long
    Dim col As Long
    col = 2 'Commence à la col 2
    Do While wsUsage.Cells(1, col).Value <> vbNullString
        feuilleNom = wsUsage.Cells(1, col).Value
        
        'Vérifier si la feuille existe dans le classeur maître
        On Error Resume Next
        Dim wsMaster As Worksheet
        Set wsMaster = wbMaster.Sheets(feuilleNom)
        On Error GoTo 0
        
        If Not wsMaster Is Nothing Then
            'Compter les lignes utilisées dans la col A
            lastUsedRow = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row
        Else
            'Si la feuille n'existe pas, assigner 0
            lastUsedRow = 0
        End If
        
        'Écrire le résultat dans la ligne disponible
        wsUsage.Cells(LigneDisponible, col).Value = lastUsedRow
        
        'Passer à la col suivante
        col = col + 1
    Loop
    
    'Fermer le classeur maître sans enregistrer
    wbMaster.Close False
    
    Application.ScreenUpdating = True
    
    'Sauvegarder et fermer le classeur d'usage
    wbUsage.Close SaveChanges:=True
    
    Call EnregistrerLogApplication("modAppli_Utils:NoterNombreLignesParFeuille", vbNullString, startTime)

End Sub

Sub ChargerPlageDansDictionary(ByRef dict As Object, ByVal rng As Range, Optional colValeurOffset As Long = 0)

    'Créer un dictionnaire si non initialisé
    If dict Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
    End If

    'Parcourir chaque cellule de la plage et ajouter au dictionnaire
    Dim cell As Range
    Dim clé As Variant
    Dim valeur As Variant
    For Each cell In rng
        clé = cell.Value
        valeur = cell.offset(0, colValeurOffset).Value 'Colonne adjacente ou selon décalage

        'Ajouter au dictionnaire si la clé n'existe pas déjà
        If Not dict.Exists(clé) Then
            dict.Add clé, valeur
        End If
    Next cell
    
End Sub

Sub ExempleUtilisation()

    Dim dict As Object
    
    'Définir la feuille de calcul et la plage
    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion.offset(2, 0)
    'Redimensionner la plage pour inclure uniquement les lignes restantes
    Set rng = rng.Resize(rng.Rows.count - 2, rng.Columns.count)

    'Charger les données dans un dictionnaire
    Call ChargerPlageDansDictionary(dict, rng, 2) ' 2 = Décalage de colonne pour les valeurs (colonne C)

    'Nettoyer la mémoire
    Set rng = Nothing
    Set ws = Nothing
    
End Sub

Sub VerifierCombinaisonClientIDClientNomDansTEC()

    'Fichier maître des clients
    Dim strFile As String
    strFile = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
    Dim wb As Workbook
    Set wb = Workbooks.Open(strFile)
    
    'Feuille TEC_Local
    Dim ws As Worksheet
    Set ws = wsdTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    'Transfère la feuille en mémoire (matrice)
    Dim m As Variant
    m = ws.Range("A3:P" & lastUsedRow).Value
    
    'Feuille de travail (ouput)
    Dim output As Worksheet
    Set output = ThisWorkbook.Sheets("Feuil1")
    Dim r As Integer
    r = 1
    output.Cells.Clear
    output.Cells(r, 1) = "Ligne"
    output.Cells(r, 2) = "TEC_ID"
    output.Cells(r, 3) = "ClientID"
    output.Cells(r, 4) = "ClientName"
    output.Cells(r, 5) = "clientNameFromMF"
    output.Cells(r, 6) = "Date"
    output.Cells(r, 7) = "Prof"
    output.Cells(r, 8) = "Description"
    output.Cells(r, 9) = "Heures"
    output.Cells(r, 10) = "estFacturée"
    
    Dim clientID As String, clientName As String, clientNameFromMF As String
    Dim allCols As Variant
    Dim i As Integer
    For i = 1 To UBound(m, 1)
        clientID = m(i, fTECClientID)
        clientName = m(i, fTECClientNom)
        
        'Obtenir le nom du client associé à clientID
        allCols = Fn_Get_A_Row_From_A_Worksheet("BD_Clients", clientID, fClntFMClientID)
        'Vérifier le résultat retourné
        If IsArray(allCols) Then
            clientNameFromMF = allCols(1)
        Else
            MsgBox "Valeur non trouvée !!!", vbCritical
        End If
        
        If clientName <> clientNameFromMF Then
            r = r + 1
            output.Cells(r, 1).Value = i + 2
            output.Cells(r, 2).Value = m(i, fTECTECID)
            output.Cells(r, 3).Value = clientID
            output.Cells(r, 4).Value = clientName
            output.Cells(r, 5).Value = clientNameFromMF
            output.Cells(r, 6).Value = m(i, fTECDate)
            output.Cells(r, 7).Value = m(i, fTECProf)
            output.Cells(r, 8).Value = m(i, fTECDescription)
            output.Cells(r, 9).Value = m(i, fTECHeures)
            output.Cells(r, 10).Value = m(i, fTECEstFacturee)
        End If
        
    Next i

    Debug.Print lastUsedRow, UBound(m, 1)

    wb.Close False
    
End Sub

Private Function AnalyserLigneTEC(data As Variant, i As Long, ByVal wsOutput As Worksheet, ByRef r As Long, _
                                  ByRef minDate As Date, ByRef maxDate As Date, _
                                  ByRef dictClient As Object, ByRef dictFacture As Object, _
                                  ByRef dictFactureHres As Object, ByRef dictDateCharge As Object, _
                                  ByRef dictTimeStamp As Object, ByRef dictDict As Object, dict_Prof As Object, _
                                  ByRef stats As StatistiquesTEC, lastTECIDReported As Long) As Boolean
                                  
    Dim isValid As Boolean
    isValid = True

    '1. On teste le code de client
    Dim codeClient As String
    codeClient = Trim$(data(i, fTECClientID))
    
    Dim tecID As Long
    tecID = data(i, fTECTECID)
    
    If Not dictClient.Exists(codeClient) Then
        Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & "', le code de client '" & codeClient & "' est INVALIDE !!!")
        isValid = False
    End If
    
    '2. On teste la date
    Dim dateTEC As Variant
    dateTEC = data(i, fTECDate)
    
    If Not IsDate(dateTEC) Or dateTEC > Date Then
        Call AjouterMessage(wsOutput, r, 2, "********** TECID =" & tecID & " a une date INVALIDE '" & dateTEC & "'")
        stats.cas_date_invalide = stats.cas_date_invalide + 1
        isValid = False
    Else
        If dateTEC <> Int(dateTEC) Then
            Call AjouterMessage(wsOutput, r, 2, "********** La date du TEC '" & dateTEC & "' n'est pas du bon format (H:M:S) pour le TECID =" & tecID)
            isValid = False
        End If
        If dateTEC > Date Then
            Call AjouterMessage(wsOutput, r, 2, "********** TECID =" & tecID & " a une date FUTURE '" & dateTEC & "'")
            stats.cas_date_future = stats.cas_date_future + 1
            isValid = False
        End If
        If dateTEC < minDate Then minDate = dateTEC
        If dateTEC > maxDate Then maxDate = dateTEC
    End If
    
    '3. Teste la valeur de 'estDetruit'
    Dim estDetruit As String
    estDetruit = UCase(Trim$(CStr(data(i, fTECEstDetruit))))
    
    If (estDetruit <> "VRAI" And estDetruit <> "FAUX") Or Len(estDetruit) <> 4 Then
        Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & "', la valeur de la colonne 'EstDetruit' est INVALIDE '" & estDetruit & "' !!!")
        stats.cas_estDetruit_invalide = stats.cas_estDetruit_invalide + 1
        isValid = False
    End If
    
    '4. Teste la valeur 'estFacturable'
    Dim estFacturable As String
    estFacturable = UCase(Trim$(CStr(data(i, fTECEstFacturable))))
    
    If InStr("VRAI^FAUX^", estFacturable & "^") = 0 Or Len(estFacturable) <> 4 Then
        Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & "', la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
        stats.cas_estFacturable_invalide = stats.cas_estFacturable_invalide + 1
        isValid = False
    End If
    
    '5. Teste la valeur de 'estFacturee'
    Dim estFacturee As String
    estFacturee = UCase(Trim$(data(i, fTECEstFacturee)))
    
    If estFacturee <> "VRAI" And estFacturee <> "FAUX" Then
        Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & "', la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
        stats.cas_estFacturee_invalide = stats.cas_estFacturee_invalide + 1
        isValid = False
    End If
    
    '6. On teste & accumule les heures
    Dim hres As Variant
    hres = data(i, fTECHeures)
    Dim prof As String
    prof = data(i, fTECProf)
    
    If Not IsNumeric(hres) Then
        Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & "', la valeur des heures est INVALIDE '" & hres & "' !!!")
        stats.cas_hres_invalide = stats.cas_hres_invalide + 1
        isValid = False
    Else
        stats.total_hres_inscrites = stats.total_hres_inscrites + CCur(hres)
    End If

    'Accumule les heures
    Dim h(1 To 6) As Currency

    'Heures INSCRITES
    h(1) = hres

    'Heures DÉTRUITES
    h(2) = 0
    If estDetruit = "VRAI" Then
        stats.total_hres_detruites = stats.total_hres_detruites + hres
        h(2) = hres
        hres = 0 'Il ne reste plus d'heures...
    End If

    'Heures FACTURABLES
    h(3) = 0
    If hres <> 0 And estFacturable = "VRAI" And Fn_Is_Client_Facturable(codeClient) = True Then
            stats.total_hres_facturables = stats.total_hres_facturables + hres
            h(3) = hres
    End If

    'Heures NON-FACTURABLES
    h(4) = 0
    If hres <> 0 Then
        stats.total_hres_non_facturables = stats.total_hres_non_facturables + hres - h(3)
        h(4) = hres - h(3)
    End If

    'Heures FACTURÉES
    h(5) = 0
    If estFacturee = "VRAI" And Fn_Is_Client_Facturable(codeClient) = True Then
            stats.total_hres_facturees = stats.total_hres_facturees + hres
            h(5) = hres
    End If

    'Heures TEC = Heures Facturables - Heures facturées
    If h(3) Then
        h(6) = h(3) - h(5)
    Else
        h(6) = 0
    End If

    If h(1) - h(2) <> h(3) + h(4) Then
        Debug.Print "#020 - " & i & " Écart - " & tecID & " " & prof & " " & dateTEC & " " & h(1) & " " & h(2) & " vs. " & h(3) & " " & h(4)
        Stop
    End If
    
    '8. Teste la date de facture, si elle existe
    Dim dateFact As Variant
    dateFact = data(i, fTECDateFacturee)
    
    If Not IsEmpty(dateFact) And CStr(dateFact) <> vbNullString Then
        If Not IsDate(dateFact) Then
            Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & "', a une date de facture INVALIDE '" & dateFact & "' !!!")
            stats.cas_date_fact_invalide = stats.cas_date_fact_invalide + 1
            isValid = False
        Else
            If dateFact > Date Then
                Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & "', a une date de facture FUTURE '" & dateFact & "' !!!")
                stats.cas_date_facture_future = stats.cas_date_facture_future + 1
                isValid = False
            End If
            If dateFact <> Int(dateFact) Then
                Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & "', la date de la facture '" & dateFact & "' n'est pas du bon format (H:M:S) !!!")
                isValid = False
            End If
        End If
    End If

    '7. On accumule les heures facturées par facture
    Dim invNo As String
    invNo = CStr(data(i, fTECNoFacture))

    If Len(invNo) > 0 Then
        If estFacturee <> "VRAI" Then
            Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & _
                "', Incongruité entre le numéro de facture '" & invNo & "' et " & _
                "'estFacturee' qui vaut '" & estFacturee & "'")
            isValid = False
        End If
    
        If invNo <> "Radiation" Then
            If Not dictFacture.Exists(invNo) Then
                Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & _
                    "', Le numéro de facture '" & invNo & "' n'existe pas dans le fichier FAC_Entête")
                isValid = False
            Else
                'Accumuler les heures dans dictFactureHres
                If dictFactureHres.Exists(invNo) Then
                    dictFactureHres(invNo) = dictFactureHres(invNo) + data(i, fTECHeures)
                Else
                    dictFactureHres.Add invNo, data(i, fTECHeures)
                End If
            End If
        End If
    Else
        If estFacturee = "VRAI" Then
            Call AjouterMessage(wsOutput, r, 2, "********** TecID = '" & tecID & _
                "', Incongruité entre le numéro de facture vide et 'estFacturee' qui vaut '" & estFacturee & "'")
            isValid = False
        End If
    End If

    '8. On alimente les 2 dictionnaires pour les dates
    Dim profID As Long
    prof = data(i, fTECProf)
    profID = data(i, fTECProfID)
    
    If CLng(tecID) > lastTECIDReported And UCase$(estDetruit) = "FAUX" Then
        Dim dCharge As Date, dStamp As Date, strCharge As String, strStamp As String
        dCharge = data(i, fTECDate)
        dStamp = data(i, fTECDateSaisie)
    
        strCharge = Format$(dCharge, "yyyy-mm-dd") & " - " & Fn_Pad_A_String(CStr(prof), " ", 5, "R")
        strStamp = Format$(dStamp, "yyyy-mm-dd") & " - " & Fn_Pad_A_String(CStr(prof), " ", 5, "R")
    
        'Accumule dans un dictionary par Date de charge
        If dictDateCharge.Exists(strCharge) Then
            dictDateCharge(strCharge) = dictDateCharge(strCharge) + CCur(hres)
        Else
            dictDateCharge.Add strCharge, CCur(hres)
        End If
    
        'Accumule dans un dictionary par TimeStamp
        If dictTimeStamp.Exists(strStamp) Then
            dictTimeStamp(strStamp) = dictTimeStamp(strStamp) + 1
        Else
            dictTimeStamp.Add strStamp, 1
        End If
    End If
    
    '10. Vérifie que chaque TecID est unique
    If dictDict.Exists(tecID) = False Then
        dictDict.Add tecID, 0
    Else
        Call AjouterMessage(wsOutput, r, 2, "********** Le TecID '" & tecID & "' est un doublon pour la ligne '" & i & "'")
        isValid = False
        stats.cas_doublon_TecID = stats.cas_doublon_TecID + 1
    End If
    
    '11. Vérifie la combinaison ProfID + Prof
    If dict_Prof.Exists(prof & "-" & profID) = False Then
        dict_Prof.Add prof & "-" & profID, 0
    End If
    
    'À la fin, retourner le résultat global pour cette ligne
    AnalyserLigneTEC = isValid
    
End Function

Public Function EgalCurrency(a As Currency, b As Currency, Optional epsilon As Currency = 0.01) As Boolean

    EgalCurrency = Abs(a - b) <= epsilon
    
End Function

Sub ExtraireCouleurCellule()

    Dim couleur As Long
    Dim rouge As Long, vert As Long, bleu As Long

    ' récupère la couleur de fond de la cellule active
    couleur = ActiveSheet.Range("B8").Interior.Color
    
    ' extrait les composantes RGB
    rouge = couleur Mod 256
    vert = (couleur \ 256) Mod 256
    bleu = (couleur \ 65536) Mod 256

    ' affiche les valeurs
    MsgBox "Rouge: " & rouge & vbCrLf & _
           "Vert: " & vert & vbCrLf & _
           "Bleu: " & bleu
End Sub
