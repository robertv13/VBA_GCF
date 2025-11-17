Attribute VB_Name = "modMenu"
Option Explicit

Sub shpMenuTEC_Click()

    Call AccederMenuTEC
    
End Sub

Sub AccederMenuTEC()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:AccederMenuTEC", vbNullString, 0)
    
    'Tous les utilisateurs ont accès au menu TEC
    Call modAppli.QuitterFeuillePourMenu(wshMenuTEC, True) '2025-08-19 @ 06:59
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU:AccederMenuTEC", vbNullString, startTime)

End Sub

Sub shpMenuFacturation_Click()

    Call AccederMenuFacturation

End Sub

Sub AccederMenuFacturation()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:AccederMenuFacturation", vbNullString, 0)
    
    If UCase(UtilisateurActif("AccesFACT")) = "VRAI" Then
        Call modAppli.QuitterFeuillePourMenu(wshMenuFAC, True) '2025-08-19 @ 07:12
    Else
        Application.EnableEvents = False
        MsgBox "Vous n'êtes pas autorisé à accéder à cette option", _
            vbInformation, _
            "Vérification des accès par utilisateur Windows"
        wshMenu.Activate
        Application.EnableEvents = True
    End If
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU:AccederMenuFacturation", vbNullString, startTime)

End Sub

Sub shpMenuComptabilite_Click()

    Call AccederMenuComptabilite
    
End Sub

Sub AccederMenuComptabilite()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:AccederMenuComptabilite", vbNullString, 0)
    
    If UCase(UtilisateurActif("AccesGL")) = "VRAI" Then
        Call modAppli.QuitterFeuillePourMenu(wshMenuGL, True) '2025-08-19 @ 07:12
    Else
        Application.EnableEvents = False
        MsgBox "Vous n'êtes pas autorisé à accéder à cette option", _
            vbInformation, _
            "Vérification des accès par utilisateur Windows"
        wshMenu.Activate
        Application.EnableEvents = True
    End If

    Call modDev_Utils.EnregistrerLogApplication("modMENU:AccederMenuComptabilite", vbNullString, startTime)

End Sub

Sub shpADMIN_Click()

    Call AccederFeuilleADMIN
    
End Sub

Sub AccederFeuilleADMIN()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:AccederFeuilleADMIN", vbNullString, 0)
    
    If modFunctions.Fn_UtilisateurWindows() = "Guillaume" Or _
            modFunctions.Fn_UtilisateurWindows() = "GuillaumeCharron" Or _
            modFunctions.Fn_UtilisateurWindows() = "gchar" Or _
            modFunctions.Fn_UtilisateurWindows() = "RobertMV" Or _
            modFunctions.Fn_UtilisateurWindows() = "robertmv" Then
        wsdADMIN.Visible = xlSheetVisible
        wsdADMIN.Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU:AccederFeuilleADMIN", vbNullString, startTime)

End Sub

Sub shpSortieApplication_Click()

    Call ConfirmerSortieApplication

End Sub

Sub ConfirmerSortieApplication()
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modMENU:ConfirmerSortieApplication", vbNullString, 0)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim shiftEnfonce As Boolean
    shiftEnfonce = (GetKeyState(vbKeyShift) < 0)
    Debug.Print "SHIFT enfoncé : " & shiftEnfonce

    If Not shiftEnfonce Then
        Dim confirmation As VbMsgBoxResult
        confirmation = MsgBox("Êtes-vous certain de vouloir quitter" & vbNewLine & vbNewLine & _
                        "l'application de gestion (sauvegarde automatique) ?", vbYesNo + vbQuestion, "Confirmation de sortie")
        If confirmation = vbYes Then
            Call modMenu.FermerApplication("Fermeture normale", False)
        End If
    Else
        Call modMenu.FermerApplication("Fermeture normale - Sauvegarde outrepassée", True)
    End If
    
End Sub

Sub FermerApplication(methode As String, ignorerSauvegarde As Boolean) '2025-09-10 @ 08:14

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modMENU:FermerApplication", "ignoreSauvegarde = '" & CStr(ignorerSauvegarde) & "'", 0)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim userName As String: userName = modFunctions.Fn_UtilisateurWindows
    
    Dim ws As Worksheet: Set ws = wsdADMIN
    
    Call ViderTableauxStructures
    
    'Effacer fichier utilisateur actif + Fermeture de la journalisation
    Call EffacerFichierUtilisateurActif(modFunctions.Fn_UtilisateurWindows())
    
    Call EnregistrerLogPerformance("Fermeture = '" & methode & "'", 0)
    
    'Fermer TOUS les formulaires (UserForm)
    Dim uf As Object
    For Each uf In VBA.UserForms
        On Error Resume Next
        Unload uf
        On Error GoTo 0
    Next
    
    'Fermer TOUS les Timer
    Call AnnulerTousLesTimers

    Call modDev_Utils.EnregistrerLogApplication("----- SESSION TERMINÉE - modMenu:FermerApplication - " & _
            methode & " -----", IIf(ignorerSauvegarde, "S A N S   S A U V E G A R D E", ""), startTime)
    
    gFermetureForcee = True
    Application.EnableEvents = False
    Dim tt0 As Double: tt0 = Timer
    If ignorerSauvegarde = False Then
        ThisWorkbook.Save
    End If
    
    Call modTraceSession.SupprimerTraceOuverture '2025-11-10 @ 08:38

    Application.DisplayAlerts = False
    If ignorerSauvegarde = False Then ThisWorkbook.Save
    ThisWorkbook.Saved = True
    
    If Workbooks.count = 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = True
    
End Sub

Sub CacherToutesFeuillesSaufMenu()
    
    DoEvents
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modMENU:CacherToutesFeuillesSaufMenu", vbNullString, 0)
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.CodeName <> "wshMenu" Then
            If modFunctions.Fn_UtilisateurWindows() <> "RobertMV" Or InStr(ws.CodeName, "wshzDoc") = 0 Then
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU:CacherToutesFeuillesSaufMenu", vbNullString, startTime)
    
End Sub

Sub CacherFormesEnFonctionUtilisateur(ByVal userName As String) '2025-06-06 @ 11:17
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modMENU:CacherFormesEnFonctionUtilisateur", vbNullString, 0)
    
    Dim ws As Worksheet: Set ws = wshMenu
    Dim devShapes As Variant
    devShapes = Array( _
        "shpImporterCorrigerMASTER", _
        "shpVérificationIntégrité", _
        "shpTraitementFichiersLog", _
        "shpSynchronisationDEVversPROD", _
        "shpAuditVBAProcedures", _
        "shpCompterLignesCode", _
        "shpRechercherCode", _
        "shpCorrigerNomClientTEC", _
        "shpCorrigerNomClientCAR", _
        "shpChercherRéférencesCirculaires", _
        "shpChangerReferenceSystem", _
        "shpListerModulesEtRoutines", _
        "shpVérificationMacrosContrôles" _
    )

    Dim isDevUser As Boolean
    isDevUser = (userName = "RobertMV" Or userName = "robertmv")
    Dim visibleState As MsoTriState
    visibleState = IIf(isDevUser, msoTrue, msoFalse)

    Dim i As Long
    For i = LBound(devShapes) To UBound(devShapes)
        On Error Resume Next 'Ignore erreur si Shape absent
        ws.Shapes(devShapes(i)).Visible = visibleState
        If Err.Number <> 0 Then
            Debug.Print "Forme introuvable: " & devShapes(i)
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    Call modDev_Utils.EnregistrerLogApplication("modMENU:CacherFormesEnFonctionUtilisateur", vbNullString, startTime)

End Sub

Sub EffacerFichierUtilisateurActif(ByVal userName As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:EffacerFichierUtilisateurActif", vbNullString, 0)
    
    Dim traceFilePath As String
    traceFilePath = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                    "SessionActive_" & userName & ".txt"
    
    If Dir(traceFilePath) <> vbNullString Then
        Kill traceFilePath
    End If
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU:EffacerFichierUtilisateurActif", vbNullString, startTime)

End Sub

Sub ViderTableauxStructures() '2025-07-01 @ 10:38

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modMENU:ViderTableauxStructures", vbNullString, 0)
    
    Dim feuilles As Variant, tableaux As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    'Feuilles & noms de tableaux à vider
    feuilles = Array("BD_Clients", "BD_Fournisseurs", "CC_Regularisations", "DEB_Recurrent", "DEB_Trans", _
                     "ENC_Details", "ENC_Entete", "FAC_Comptes_Clients", "FAC_Details", "FAC_Entete", _
                     "FAC_Projets_Details", "FAC_Projets_Entete", "FAC_Sommaire_Taux", "GL_Trans", _
                     "TEC_Local")
    tableaux = Array("l_tbl_BD_Clients", "l_tbl_Fournisseur_FM", "l_tbl_CC_Regularisations", _
                     "l_tbl_DEB_Recurrent", "l_tbl_DEB_Trans", "l_tbl_ENC_Details", "l_tbl_ENC_Entete", _
                     "l_tbl_FAC_Comptes_Clients", "l_tbl_FAC_Details", "l_tbl_FAC_Entete", _
                     "l_tbl_FAC_Projets_Details", "l_tbl_FAC_Projets_Entete", "l_tbl_FAC_Sommaire_Taux", _
                     "l_tbl_GL_Trans", "l_tbl_TEC_Local")

'    On Error Resume Next '2025-11-14 @ 18:29

    Dim i As Long
    For i = LBound(feuilles) To UBound(feuilles)
        Set ws = ThisWorkbook.Sheets(Trim$(feuilles(i)))
        Set lo = ws.ListObjects(tableaux(i))

        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                'Désactiver les filtres s'ils sont actifs '2025-11-11 @ 06:37
                If Not lo.AutoFilter Is Nothing Then
                    If lo.AutoFilter.FilterMode Then
                        lo.AutoFilter.ShowAllData
                    End If
                End If
                lo.DataBodyRange.Delete
            End If
        Else
            Debug.Print "Tableau '" & tableaux(i) & "' est introuvable dans '" & Trim(feuilles(i)) & "'"
        End If
    Next i

    On Error GoTo 0

    Call modDev_Utils.EnregistrerLogApplication("modMENU:ViderTableauxStructures", vbNullString, startTime)

End Sub

Public Sub AnnulerTousLesTimers() '2025-11-10 @ 06:32

    Debug.Print String(40, "-")
    Debug.Print "Annulation des timers OnTime en cours..."
    Debug.Print String(40, "-")

    If gNextBackupTime <> 0 Then Call AnnulerTimer("DemarrerSauvegardeCodeVBAAutomatique", gNextBackupTime)
    If gHeureProchaineVerification <> 0 Then Call AnnulerTimer("modSurveillance.VerifierActivite", gHeureProchaineVerification)
    If gProchainTick <> 0 Then Call AnnulerTimer("modSurveillance.SurveillerFermetureAuto", gProchainTick)
    If gProchainRafraichir <> 0 Then Call AnnulerTimer("ufConfirmationFermeture.RafraichirTimer", gProchainRafraichir)

    Debug.Print String(40, "-")
    Debug.Print "Fin de l’annulation des timers"
    Debug.Print String(40, "-")
    
End Sub

Public Sub AnnulerTimer(nomProcedure As String, heurePlanifiee As Date) '2025-11-10 @ 06:28

    On Error Resume Next
    
    Application.OnTime heurePlanifiee, nomProcedure, , False
    If Err.Number = 0 Then
        Debug.Print Left(nomProcedure & Space(36), 36) & " - Timer annulé"
    Else
        Debug.Print Left(nomProcedure & Space(36), 36) & " - É C H E C annulation - " & Err.Number & " - " & Err.description
    End If
    
    On Error GoTo 0
    
End Sub

Sub shpImporterCorrigerMASTER_Click()

    If modFunctions.Fn_UtilisateurWindows() <> "RobertMV" And modFunctions.Fn_UtilisateurWindows() <> "robertmv" Then
        Exit Sub
    End If
    
    'Crée un répertoire local et importe les fichiers à analyser
    Call CreerRepertoireEtImporterFichiers
    
End Sub

Sub shpVerificationIntegrite_Click()

    Call modVerifications.VerifierIntegriteTablesLocales

End Sub

Sub shpRechercherCode_Click()

    Call modDev_Utils.RechercherCodeProjet

End Sub

Sub shpCompterLignesCodeProjet_Click()

    Call modDev_Tools.CompterLignesCode

End Sub

Sub shpChercherReferencesCirculaires_Click()

    Call modDev_Tools.DetecterReferenceCirculaireDansClasseur
    
End Sub

Sub shpChangerReferenceSystem_Click()

    Call modDev_Utils.ChangerSystemeReferenceCellules
    
End Sub

Sub shpListerModulesEtRoutines_Click()

    Call modDev_Utils.ListerToutesProceduresEtFonctions
    
End Sub

Sub shpVerificationMacrosControles_Click()

    Call modAuditVBA.zz_VerifierControlesAssociesToutesFeuilles

End Sub

Sub shpRetournerMenuPrincipal_Click()

    Call RetournerMenuPrincipal

End Sub

Sub RetournerMenuPrincipal()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:RetournerMenuPrincipal", ActiveSheet.Name, 0)
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "Menu" Then ws.Visible = xlSheetHidden
    Next ws
    
    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
        .Activate
        .Range("A1").Select
    End With

    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU:RetournerMenuPrincipal", vbNullString, startTime)
    
End Sub

