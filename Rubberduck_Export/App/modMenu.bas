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

    Call SauvegarderEtSortirApplication

End Sub

Sub SauvegarderEtSortirApplication() '2024-08-30 @ 07:37
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:SauvegarderEtSortirApplication", vbNullString, 0)
    
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
            Call FermerApplicationNormalement(modFunctions.Fn_UtilisateurWindows(), "Fermeture normale")
        End If
    Else
        Call FermerApplicationNormalement(modFunctions.Fn_UtilisateurWindows(), "Sauvegarde outrepassée")
    End If
    
End Sub

Sub FermerApplicationNormalement(ByVal userName As String, flag As String, Optional ByVal ignorerSauvegarde As Boolean = False) '2025-09-10 @ 08:14

    Call modDev_Utils.EnregistrerLogApplication("modMENU:FermerApplicationNormalement - " & flag, vbNullString, -1)
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:FermerApplicationNormalement", vbNullString, startTime)
    
    On Error GoTo ExitPoint
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Application.StatusBar = False
    
    Dim ws As Worksheet
    Set ws = wsdADMIN
    
    Call ViderTableauxStructures
    
    'Effacer fichier utilisateur actif + Fermeture de la journalisation
    Call EffacerFichierUtilisateurActif(modFunctions.Fn_UtilisateurWindows())
    
    Call modDev_Utils.EnregistrerLogApplication("----- Session terminée NORMALEMENT (modMenu:SauvegarderEtSortirApplication) -----", _
        IIf(ignorerSauvegarde, "S A N S   S A U V E G A R D E", ""), 0)
        
    Call modDev_Utils.EnregistrerLogApplication(vbNullString, vbNullString, -1)

    'Fermer la vérification d'inactivité
    If gProchaineVerification > 0 Then
        On Error Resume Next
        Application.OnTime gProchaineVerification, "VerifierDerniereActivite", , False
        On Error GoTo 0
    End If

    'Fermer la sauvegarde automatique du code VBA (seul le développeur déclenche la sauvegarde automtique)
    If userName = "RobertMV" Or userName = "robertmv" Then
        Call ArreterSauvegardeCodeVBA
        Call ExporterCodeVBA
    End If
    
'    'Log de fermeture (fichier autre que log normal)
'    Open Environ("TEMP") & "\APP_FermetureEnCours.txt" For Output As #1
'    Print #1, "Session fermee volontairement"
'    Close #1
'
    'Fermeture du classeur de l'application uniquement
    If ignorerSauvegarde Then
        'Pas de sauvegarde
        Debug.Print Now() & " Fermeture du classeur - SANS SAUVEGARDE"
        ThisWorkbook.Close SaveChanges:=False
    Else
        'Avec sauvegarde
        Debug.Print Now() & " Fermeture du classeur - Avec sauvegarde"
        ThisWorkbook.Close SaveChanges:=True
    End If
    
ExitPoint:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Set ws = Nothing
    On Error GoTo 0
    
End Sub

Sub CacherToutesFeuillesSaufMenu() '2024-02-20 @ 07:28
    
    DoEvents
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:CacherToutesFeuillesSaufMenu", vbNullString, 0)
    
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
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:CacherFormesEnFonctionUtilisateur", vbNullString, 0)
    
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
    traceFilePath = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & "Actif_" & userName & ".txt"
    
    If Dir(traceFilePath) <> vbNullString Then
        Kill traceFilePath
    End If
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU:EffacerFichierUtilisateurActif", vbNullString, startTime)

End Sub

Sub ViderTableauxStructures() '2025-07-01 @ 10:38

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU:ViderTableauxStructures", vbNullString, 0)
    
    Dim feuilles As Variant, tableaux As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    'Feuilles & noms de tableaux à vider
    feuilles = Array("BD_Clients", "BD_Fournisseurs", _
                     "CC_Regularisations", _
                     "DEB_Trans", _
                     "ENC_Details", "ENC_Entete", _
                     "FAC_Comptes_Clients", "FAC_Details", "FAC_Entete", _
                     "FAC_Projets_Details", "FAC_Projets_Entete", _
                     "FAC_Sommaire_Taux", _
                     "GL_Trans", _
                     "TEC_Local")
    tableaux = Array("l_tbl_BD_Clients", "l_tbl_Fournisseur_FM", _
                     "l_tbl_CC_Regularisations", _
                     "l_tbl_DEB_Trans", _
                     "l_tbl_ENC_Details", "l_tbl_ENC_Entete", _
                     "l_tbl_FAC_Comptes_Clients", "l_tbl_FAC_Details", "l_tbl_FAC_Entete", _
                     "l_tbl_FAC_Projets_Details", "l_tbl_FAC_Projets_Entete", _
                     "l_tbl_FAC_Sommaire_Taux", _
                     "l_tbl_GL_Trans", _
                     "l_tbl_TEC_Local")

    On Error Resume Next

    Dim i As Long
    For i = LBound(feuilles) To UBound(feuilles)
        Set ws = ThisWorkbook.Sheets(Trim$(feuilles(i)))
        Set lo = ws.ListObjects(tableaux(i))

        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                lo.DataBodyRange.Delete
            End If
        Else
            Debug.Print "Tableau '" & tableaux(i) & "' est introuvable dans '" & Trim(feuilles(i)) & "'"
        End If
    Next i

    On Error GoTo 0

    Call modDev_Utils.EnregistrerLogApplication("modMENU:ViderTableauxStructures", vbNullString, startTime)

End Sub

Sub shpImporterCorrigerMASTER_Click()

    If modFunctions.Fn_UtilisateurWindows() <> "RobertMV" And modFunctions.Fn_UtilisateurWindows() <> "robertmv" Then
        Exit Sub
    End If
    
    'Crée un répertoire local et importe les fichiers à analyser
    Call CreerRepertoireEtImporterFichiers
    
End Sub

Sub shpVerificationIntegrite_Click()

    Call modAppli_Utils.VerifierIntegriteTablesLocales

End Sub

Sub shpRechercherCode_Click()

    Call modDev_Utils.RechercherCodeProjet

End Sub

Sub shpCompterLignesCodeProjet_Click()

    Call CompterLignesCode

End Sub

Sub shpCorrigerNomClientTEC_Click()

    Call modzDataConversion.CorrigerNomClientDansTEC
    
End Sub

Sub shpCorrigerNomClientCAR_Click()

    Call modzDataConversion.CorrigerNomClientDansCAR
    
End Sub

Sub shpChercherReferencesCirculaires_Click() '2024-11-22 @ 13:33

    Call modDev_Tools.DetecterReferenceCirculaireDansClasseur
    
End Sub

Sub shpChangerReferenceSystem_Click() '2024-11-22 @ 13:33

    Call modDev_Utils.ChangerSystemeReferenceCellules
    
End Sub

Sub shpListerModulesEtRoutines_Click() '2024-11-22 @ 13:33

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

