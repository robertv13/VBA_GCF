Attribute VB_Name = "modAppli"
Option Explicit

Public gProchaineVerifUserForm As Date
Public gHeurePrevueFermetureAutomatique As Date 'Heure à laquelle l'application devrait fermer
Public gProchainTick As Date                 'Heure du compte à rebours
Public gClignoteEtat As Boolean

Sub DemarrerApplication(uw As String) '2025-07-11 @ 15:16

    'Mise en place du répertoire de base (C:\... ou P:\...)
    Dim rootPath As String
    rootPath = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
    Application.EnableEvents = False
    wsdADMIN.Range("PATH_DATA_FILES").Value = rootPath
    Application.EnableEvents = True
   
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("----- DÉBUT D'UNE NOUVELLE SESSION (modAppli:DemarrerApplication) -----", vbNullString, 0)
    
    'Initialisation de la session utilisateur '2025-10-19 @ 11:24
    Call InitialiserSessionUtilisateur
    
    On Error GoTo ErrorHandler
    
    If Application.EnableEvents = False Then Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.StatusBar = "Vérification de l'accès au répertoire principal"
    If Fn_AccesServeur(rootPath) = False Then
        MsgBox "Le répertoire principal '" & rootPath & "' n'est pas accessible." & vbNewLine & vbNewLine & _
               "Veuillez vérifier votre connexion au serveur SVP", vbCritical, rootPath
        Exit Sub
    End If
    Application.StatusBar = False

    Call CreerFichierUtilisateurActif(uw)
    Call FixerFormatDateUtilisateur(uw)
    Call CreerSauvegardeMaster
    Call EcrireInformationsConfigAuMenu
    wshMenu.Range("A1").Value = wsdADMIN.Range("NomEntreprise").Value
    Call modMenu.CacherFormesEnFonctionUtilisateur(uw)
    
    'Protection de la feuille wshMenu
    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
    'Efface les feuilles dont le codename n'est pas wsh* -ET- dont le nom commence par 'Feuil'
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If InStr(ws.CodeName, "wsh") <> 1 And InStr(ws.CodeName, "Feuil") = 1 Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    wshMenu.Activate

    If UtilisateurActif("Role") = "Dev" Then
        Call DemarrerSauvegardeCodeVBAAutomatique
    End If
    
    'Libérer la mémoire
    Set wb = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli:DemarrerApplication", vbNullString, startTime)
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True 'On s'assure de toujours restaurer l'état
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Call modDev_Utils.EnregistrerLogApplication("modAppli:DemarrerApplication (ERREUR) : " & Err.description, Timer)
    
End Sub

Sub VerifierVersionApplication(path As String, versionApplication As String) '2025-08-12 @ 16:08

    Dim versionData As String
    Dim utilisateurWindows As String
    On Error GoTo ErreurLecture
    versionData = Trim(Fn_LireFichierTXT(path & Application.PathSeparator & "APP_Version.txt"))
    
    If versionData <> versionApplication And _
        modFunctions.Fn_UtilisateurWindows() <> "RobertMV" And _
        modFunctions.Fn_UtilisateurWindows() <> "robertmv" Then
        MsgBox "La version de l'application (" & versionApplication & ") ne correspond pas" & vbCrLf & vbCrLf & _
               "à la version des données (" & versionData & ")." & vbCrLf & vbCrLf & _
               "Veuillez mettre à jour votre application -OU-" & vbCrLf & vbCrLf & _
               "Contactez le développeur", _
               vbCritical, _
               "Version de l'application incompatible avec les données"
               
        Call FermerApplicationNormalement(modFunctions.Fn_UtilisateurWindows(), "Erreur de Version")
    End If
    Exit Sub

ErreurLecture:
    MsgBox "Impossible de lire le fichier de version du répertoire" & vbNewLine & vbNewLine & _
            path, _
            vbExclamation, _
            "Impossible de lire la version des données"
    
    Call FermerApplicationNormalement(modFunctions.Fn_UtilisateurWindows(), "Impossible de comparer la Version")
    
End Sub

Function Fn_RepertoireBaseApplication(uw As String) As String '2025-03-03 @ 20:28
   
    DoEvents
    
    If uw = "RobertMV" Or uw = "robertmv" Then
        Fn_RepertoireBaseApplication = "C:\VBA\GC_FISCALITÉ"
    Else
        Fn_RepertoireBaseApplication = "P:\Administration\APP\GCF"
    End If

End Function

Sub CreerFichierUtilisateurActif(ByVal userName As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli:CreerFichierUtilisateurActif", vbNullString, 0)
    
    Dim traceFilePath As String
    traceFilePath = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & "Actif_" & userName & ".txt"
    
    Dim FileNumber As Long
    FileNumber = FreeFile
    
    On Error GoTo Error_Handling
    Open traceFilePath For Output As FileNumber
    On Error GoTo 0
    
    Print #FileNumber, "Utilisateur " & userName & " a ouvert l'application à " & Format$(Now(), "yyyy-mm-dd hh:mm:ss") & " - Version " & ThisWorkbook.Name
    Close FileNumber
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli:CreerFichierUtilisateurActif", vbNullString, startTime)
    
    Exit Sub

Error_Handling:
    MsgBox _
        Prompt:="Erreur en tentant d'accéder le répertoire" & vbNewLine & vbNewLine & _
                    "'" & traceFilePath & "'" & vbNewLine & vbNewLine & _
                    "Erreur # " & Err.Number & " - " & Err.description, _
        Title:="Accès à " & traceFilePath, _
        Buttons:=vbCritical

End Sub

Sub FixerFormatDateUtilisateur(ByVal user As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli:FixerFormatDateUtilisateur", vbNullString, 0)

    Dim userDateFormat As String
    userDateFormat = UtilisateurActif("FormatDate")
    If userDateFormat = vbNullString Then
        userDateFormat = "dd/mm/yyyy"
    End If
    
    wsdADMIN.Range("B1").Value = userDateFormat
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli:FixerFormatDateUtilisateur", vbNullString, startTime)
    
End Sub

Sub CreerSauvegardeMaster()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli:CreerSauvegardeMaster", vbNullString, 0)
    
    On Error GoTo MASTER_NOT_AVAILABLE
    
    'Chemin source (fichier principal) et destination (sauvegarde)
    Dim masterFileFullPath As String
    Dim masterFileName As String
    
    masterFileFullPath = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & wsdADMIN.Range("MASTER_FILE").Value
    masterFileName = wsdADMIN.Range("MASTER_FILE").Value
    If Right(masterFileName, 5) = ".xlsx" Then
        masterFileName = Left(masterFileName, Len(masterFileName) - 5)
    End If
    
    Dim backupFileFullPath As String
    backupFileFullPath = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                         masterFileName & "_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    
    'Créer directement une copie du fichier sans ouvrir Excel
    FileCopy masterFileFullPath, backupFileFullPath

    Call modDev_Utils.EnregistrerLogApplication("modAppli:CreerSauvegardeMaster", vbNullString, startTime)
    
    Exit Sub
    
MASTER_NOT_AVAILABLE:
    MsgBox _
        Prompt:="Le fichier GCF_MASTER.xlsx ne peut être accédé..." & vbNewLine & vbNewLine & _
                    "Le fichier nécessite une réparation manuelle", _
        Title:="Situation anormale (" & Err.Number & " " & Err.description & ")", _
        Buttons:=vbCritical
    Application.Quit

End Sub

Sub EcrireInformationsConfigAuMenu()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli:EcrireInformationsConfigAuMenu", vbNullString, 0)
    
    Dim oldEnableEvents As Boolean

    Dim environnement As String
    Dim formatDate As String

    Dim valeurs As Variant
    
    oldEnableEvents = Application.EnableEvents
    On Error GoTo CleanUp

    wshMenu.Unprotect

    Application.EnableEvents = False

    ' Récupération des valeurs
    formatDate = wsdADMIN.Range("B1").Value
    environnement = wsdADMIN.Range("PATH_DATA_FILES").Value

    valeurs = Array( _
        "Heure - " & Format$(Now(), formatDate & " hh:mm:ss"), _
        "Version - " & ThisWorkbook.Name, _
        "Utilisateur - " & UtilisateurActif("Prenom"), _
        "Environnement - " & environnement, _
        "Format de la date - " & formatDate)

    ' Ecriture en une seule opération
    wshMenu.Range("A30:A34").Value = Application.WorksheetFunction.Transpose(valeurs)

    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

CleanUp:
    Application.EnableEvents = oldEnableEvents
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli:EcrireInformationsConfigAuMenu", vbNullString, startTime)
    
End Sub

Public Sub ConnecterControlesDeForme(frm As Object) '2025-05-30 @ 13:22

    Set colWrappers = New Collection
    Call ConnecterControlesDeFormeRecursivement(frm.Controls)
    
End Sub

Private Sub ConnecterControlesDeFormeRecursivement(ctrls As MSForms.Controls) '2025-05-30 @ 13:22

    Dim ctrl As MSForms.Control
    For Each ctrl In ctrls
        If TypeName(ctrl) <> "Label" Then
            Select Case TypeName(ctrl)
                Case "Frame", "TabStrip"
                    Call ConnecterControlesDeFormeRecursivement(ctrl.Controls)
                Case "MultiPage"
                    Dim i As Integer
                    For i = 0 To ctrl.Pages.count - 1
                        Call ConnecterControlesDeFormeRecursivement(ctrl.Pages(i).Controls)
                    Next i
                Case Else
                    On Error Resume Next
                    Dim wrapper As New clsControlWrapper
                    Set wrapper.ctrl = ctrl
                    colWrappers.Add wrapper, ctrl.Name
                    On Error GoTo 0
            End Select
        End If
    Next ctrl
    
End Sub

Public Sub EnregistrerActivite(Optional ByVal msg As String = vbNullString) '2025-07-02 @ 15:19
    
    If TimeValue(Now) < TimeSerial(gHEURE_DEBUT_SURVEILLANCE, 0, 0) Then
        Exit Sub
    End If

    'Noter état de EnableEvents
    Dim activeEvents As Boolean
    activeEvents = Application.EnableEvents
    If activeEvents = True Then Application.EnableEvents = False

    'Enregistrer la dernière activité
    gDerniereActivite = Now
    Call EnregistrerActiviteAuLog(msg) '2025-07-03 @ 10:31
    
    'Rétablir l'état de EnableEvents
    If activeEvents <> Application.EnableEvents Then
        Application.EnableEvents = activeEvents
    End If

End Sub

'@Description ("Vérifie la dernière activité et lance fermeture si plus de x minutes")
Public Sub VerifierDerniereActivite() '2025-07-02 @ 12:10
Attribute VerifierDerniereActivite.VB_Description = "Vérifie l'inactivité et ferme si plus de x minutes"

    On Error GoTo GestionErreur

    'Ne rien faire avant l'heure de début de la surveillance
    If TimeValue(Now) < TimeSerial(gHEURE_DEBUT_SURVEILLANCE, 0, 0) Then
        Call PlanifierVerificationDerniereActivite
        Exit Sub
    End If

    'Vérification de l'initialisation
    If gDerniereActivite = 0 Then
        If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:VerifierDerniereActivite] gDerniereActivite non initialisée"
        Exit Sub
    End If

    'Calcul du temps d'inactivité en minutes
    Dim minutesInactives As Double
    minutesInactives = Round(Fn_MinutesDepuisDerniereActivite(), 1)

    If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:VerifierDerniereActivite] Inactif depuis " & minutesInactives & " min. - " & _
                                    "Fréq. vérification = "; gFREQUENCE_VERIFICATION_INACTIVITE & " min., " & _
                                    "Durée max. sans activité = " & gMAXIMUM_MINUTES_INACTIVITE & " min., " & _
                                    "Délai de grâce (dernière chance) = " & gDELAI_GRACE_SECONDES & " sec."
    
    'Barre d’état informative
    Dim minute1 As String
    Dim minute2 As String
    'Minute ou minutes (minute1)
    If minutesInactives <= 1 Then
        minute1 = "minute"
    Else
        minute1 = "minutes"
    End If
    'Minute ou minutes (minute2)
    If gMAXIMUM_MINUTES_INACTIVITE - minutesInactives <= 1 Then
        minute2 = "minute"
    Else
        minute2 = "minutes"
    End If
    If minutesInactives < gMAXIMUM_MINUTES_INACTIVITE Then
        Application.StatusBar = "Aucune activité dans l'application depuis " & _
            Format$(minutesInactives, "0") & " " & minute1 & " - Fermeture planifiée dans " & _
            Format$(gMAXIMUM_MINUTES_INACTIVITE - minutesInactives, "0") & " " & minute2 & " - " & _
            Format$(Now, "hh:mm:ss")
    Else
        Application.StatusBar = False
    End If

    'Fermeture si délai dépassé, on passe à la dernier chance...
    If minutesInactives >= gMAXIMUM_MINUTES_INACTIVITE Then
        If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:VerifierDerniereActivite] Inactivité trop longue (" & Format$(minutesInactives, "0") & " minutes) — fermeture de l'application"
        gFermeturePlanifiee = GetProchaineFermeture()
        If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:VerifierDerniereActivite] Avant l'ajout de " & gDELAI_GRACE_SECONDES & " secondes, gFermeture = " & Format(gFermeturePlanifiee, "; hh: mm: ss "); vbNullString
        gFermeturePlanifiee = Now + TimeSerial(0, 0, gDELAI_GRACE_SECONDES)
        If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:VerifierDerniereActivite] Après l'ajout de " & gDELAI_GRACE_SECONDES & " secondes, gFermeture = " & Format(gFermeturePlanifiee, "hh:mm:ss")
        
        On Error Resume Next
        If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:VerifierDerniereActivite] OnTime prévu pour : " & Format(gFermeturePlanifiee, "hh:mm:ss")
        Application.OnTime gFermeturePlanifiee, "FermerApplicationInactive"
        On Error GoTo 0
        
        Unload ufConfirmationFermeture '2025-07-02 @ 07:54
        
        Call ufConfirmationFermeture.AfficherMessage(minutesInactives)
        Exit Sub
    End If

    'Replanification
    Call PlanifierVerificationDerniereActivite
    Exit Sub

GestionErreur:
    Debug.Print "[modAppli:VerifierDerniereActivite] Erreur dans VerifierDerniereActivite : " & Err.Number & " - " & Err.description

End Sub

Public Sub PlanifierVerificationDerniereActivite() '2025-07-01 @ 13:53
    
    'Ne rien faire avant l'heure de début de la surveillance '2025-10-31 @08:24
    If TimeValue(Now) < TimeSerial(gHEURE_DEBUT_SURVEILLANCE, 0, 0) Then
        Call PlanifierVerificationDerniereActivite
        Exit Sub
    End If
    
    On Error Resume Next
    Application.OnTime gProchaineVerification, "VerifierDerniereActivite", , False
    On Error GoTo 0
    
    gProchaineVerification = Now + TimeSerial(0, gFREQUENCE_VERIFICATION_INACTIVITE, 0)
    Application.OnTime gProchaineVerification, "VerifierDerniereActivite"
    If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:PlanifierVerificationDerniereActivite] Prochaine vérification est prévue à " & Format(gProchaineVerification, "hh:mm:ss")
    
End Sub

Public Sub FermerApplicationInactive() '2025-07-02 @ 06:19

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli:FermerApplicationInactive", vbNullString, 0)
    
    'Ajoute un log pour vérification
    If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:FermerApplicationInactive] Fermeture AUTOMATIQUE déclenchée à : " & Format(Now, "hh:mm:ss")

    Call modDev_Utils.EnregistrerLogApplication("modAppli:FermerApplicationInactive", vbNullString, startTime)
    
    'Appel direct de la procédure de fermeture
    Call FermerApplicationNormalement(modFunctions.Fn_UtilisateurWindows(), "Application est Inactive")
    
End Sub

Public Sub RelancerTimer() '2025-07-02 @ 06:43

    If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:RelancerTimer] Appel de 'ufConfirmationFermeture.RafraichirTimer'"
    ufConfirmationFermeture.RafraichirTimer
    
End Sub

Public Sub EnregistrerActiviteAuLog(ByVal message As String) '2025-10-30 @ 07:44

    'Ne rien faire avant l'heure de début de la surveillance
    If TimeValue(Now) < TimeSerial(gHEURE_DEBUT_SURVEILLANCE, 0, 0) Then
        Exit Sub
    End If

    Dim cheminLog As String
    cheminLog = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & "\ActiviteDurantSurveillance.txt"

    Dim fileNum As Integer
    fileNum = FreeFile

    Open cheminLog For Append As #fileNum
    Print #fileNum, "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] [" & ThisWorkbook.Name & "] [" & modFunctions.Fn_UtilisateurWindows & "] [" & _
                        Fn_ContexteActifComplet() & "] [" & message & "]"
    Close #fileNum
    
End Sub

Sub LancerSurveillanceUserForm() '2025-08-29 @ 18:32

    gProchaineVerifUserForm = Now + TimeSerial(0, 0, 10) ' toutes les 10 secondes
    Application.OnTime gProchaineVerifUserForm, "VerifierInactiviteUserForm"
    
End Sub

Sub AnnulerSurveillanceUserForm() '2025-08-29 @ 18:32

    On Error Resume Next
    Application.OnTime gProchaineVerifUserForm, "VerifierInactiviteUserForm", , False
    
End Sub

Sub VerifierInactiviteUserForm() '2025-08-29 @ 18:32

    If Fn_MinutesDepuisDerniereActivite() >= gMAXIMUM_MINUTES_INACTIVITE Then
        If gMODE_DEBUG Then Debug.Print Now() & " [modAppli:VerifierInactiviteUserForm] Inactivité détectée dans UserForm — fermeture"
        Unload ufSaisieHeures
        Call FermerApplicationInactive
    Else
        Call LancerSurveillanceUserForm
    End If
    
End Sub

Sub QuitterFeuillePourMenu(ByVal nomFeuilleMenu As Worksheet, Optional masquerFeuilleActive As Boolean = False) '2025-08-19 @ 06:46

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If masquerFeuilleActive And ActiveSheet.Name <> "Menu" Then ActiveSheet.Visible = xlSheetHidden

    nomFeuilleMenu.Visible = xlSheetVisible
    nomFeuilleMenu.Activate
    nomFeuilleMenu.Range("A1").Select

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub
    
Sub AfficherErreurCritique(message As String) '2025-10-19 @ 10:36

    MsgBox message, _
        vbCritical, _
        "Erreur critique dans l'application"
    
End Sub

