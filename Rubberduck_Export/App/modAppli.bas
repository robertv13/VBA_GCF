Attribute VB_Name = "modAppli"
Option Explicit

Public gProchaineVerifUserForm As Date
Public gHeurePrevueFermetureAutomatique As Date 'Heure à laquelle l'application devrait fermer
Public gClignoteEtat As Boolean

Sub DemarrerApplication(userWindows As String) '2025-07-11 @ 15:16

    'Mise en place du répertoire de base (C:\... ou P:\...)
    Dim rootPath As String
    rootPath = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
    Application.EnableEvents = False
    wsdADMIN.Range("PATH_DATA_FILES").Value = rootPath
    Application.EnableEvents = True
   
    Dim startTime As Double: startTime = Timer:
    Call modDev_Utils.EnregistrerLogApplication( _
                    "----- DÉBUT D'UNE NOUVELLE SESSION (modAppli:DemarrerApplication) -----", _
                    vbNullString, 0)
    
    'Initialisation de la session utilisateur '2025-10-19 @ 11:24
    Call InitialiserSessionUtilisateur
    
'    On Error GoTo ErrorHandler
    
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

    Call CreerFichierUtilisateurActif(userWindows)
    Call FixerFormatDateUtilisateur(userWindows)
    Call CreerSauvegardeMaster
    Call EcrireInformationsConfigAuMenu
    
    wshMenu.Range("A1").Value = wsdADMIN.Range("NomEntreprise").Value
    
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
    
    If UtilisateurActif("Role") = "Dev" Then
        Call DemarrerSauvegardeCodeVBAAutomatique
    End If
    
    Set ws = wshMenu
    Call InitialiserFeuilleMenu(ws, "modAppli:DemarrerApplication", "Menu Principal - 1") '2025-11-10 @ 08:51

    'Libérer la mémoire
    Set wb = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli:DemarrerApplication", vbNullString, startTime)
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True 'On s'assure de toujours restaurer l'état
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Call EnregistrerErreurs("modAppli", "DemarrerApplication", "Au démarrage", Err.Number, "ERREUR")
    Call modDev_Utils.EnregistrerLogApplication("modAppli:DemarrerApplication (ERREUR) : " & Err.description, "", startTime)
    
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
               
        Call modMenu.FermerApplication("Incompatibilité de numéro de version", True)
    End If
    Exit Sub

ErreurLecture:
    MsgBox "Impossible de lire le fichier de version du répertoire" & vbNewLine & vbNewLine & _
            path, _
            vbExclamation, _
            "Impossible de lire la version des données"
    
    Call modMenu.FermerApplication("Incapable de vérifier numéro de version", True)
    
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

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modAppli:CreerFichierUtilisateurActif", vbNullString, 0)
    
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
    Dim t0 As Double: t0 = Timer
    FileCopy masterFileFullPath, backupFileFullPath

    Call modDev_Utils.EnregistrerLogApplication("modAppli:CreerSauvegardeMaster", vbNullString, startTime)
    
    Exit Sub
    
MASTER_NOT_AVAILABLE:
    
    Call EnregistrerErreurs("modAppli", "CreerSauvegardeMaster", "Fichier MASTER n'est pas disponible - ", _
            Err.Number, "ERREUR")
    MsgBox _
        Prompt:="Le fichier GCF_MASTER.xlsx ne peut être accédé..." & vbNewLine & vbNewLine & _
                    "Le fichier nécessite une réparation manuelle", _
        Title:="Situation anormale (" & Err.Number & " " & Err.description & ")", _
        Buttons:=vbCritical
    ThisWorkbook.Close SaveChanges:=False

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

Sub QuitterFeuillePourMenu(ByVal nomFeuilleMenu As Worksheet, Optional masquerFeuilleActive As Boolean = False) '2025-08-19 @ 06:46

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modAppli:QuitterFeuillePourMenu", vbNullString, 0)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If masquerFeuilleActive And ActiveSheet.Name <> "Menu" Then ActiveSheet.Visible = xlSheetHidden

    nomFeuilleMenu.Visible = xlSheetVisible
    nomFeuilleMenu.Activate
    nomFeuilleMenu.Range("A1").Select

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli:QuitterFeuillePourMenu", vbNullString, startTime)

End Sub
    
Sub AfficherErreurCritique(modApp As String, procName As String, message As String) '2025-11-05 @ 07:39

    Call EnregistrerErreurs(modApp, procName, message, 0, "CRITICAL")
    MsgBox message, _
        vbCritical, _
        "Erreur critique dans l'application"
    
End Sub

Public Sub EnregistrerLogPerformance(nomProcedure As String, duree As Double) '2025-10-31 @ 14:05

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modAppli:EnregistrerLogPerformance", vbNullString, 0)
    
    On Error Resume Next

    'Définir le chemin du fichier log (local ou partagé)
    Dim fullPathPerformanceLog As String
    fullPathPerformanceLog = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & _
                                Application.PathSeparator & "Performance.log"
    
    'Obtenir l'utilisateur Windows
    Dim utilisateur As String
    utilisateur = Left$(Trim(Environ("USERNAME")) & Space(16), 16)

    'Horodatage complet
    Dim horodatage As String
    horodatage = Format(Now, "yyyy-mm-dd hh:nn:ss")

    'Construire la ligne de log
    Dim ligneLog As String
    Select Case duree
        Case Is > 0
            ligneLog = horodatage & " | " & utilisateur & " | " & ThisWorkbook.Name & " | " & _
                                            nomProcedure & " | " & Format(duree, "0.000") & " sec"
        Case Is = 0
            ligneLog = horodatage & " | " & utilisateur & " | " & ThisWorkbook.Name & " | " & _
                                            nomProcedure
        Case Is < 0
            ligneLog = vbNullString
        End Select

    Dim canalLog As Integer
    canalLog = FreeFile
    
    Open fullPathPerformanceLog For Append As #canalLog
    
    'Écrire dans le fichier
    Print #canalLog, ligneLog
    
    Close #canalLog

    On Error GoTo 0
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli:EnregistrerLogPerformance", vbNullString, startTime)

End Sub

Public Sub EnregistrerErreurs(moduleAppelant As String, _
                                nomProcedure As String, _
                                commentaire As String, _
                                Optional numeroErreur As Variant = 0, _
                                Optional niveauGravite As String = "ERREUR") '2025-11-05 @ 07:16
                       
    Dim horodatage As String: horodatage = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    Dim description As String
    If IsNumeric(numeroErreur) And numeroErreur <> 0 Then
        description = Err.description
    Else
        description = commentaire ' commentaire métier si pas d’erreur système
    End If
    
    Dim utilisateur As String
    utilisateur = Environ("USERNAME")
    
    Dim ligneLog As String
    ligneLog = horodatage & " | " & utilisateur & " | " & ActiveWorkbook.Name & " | " & moduleAppelant & _
                "." & nomProcedure & " | " & niveauGravite & " | " & _
                CStr(numeroErreur) & " | " & description
               
    Dim fullPathErreurLog As String
    
    fullPathErreurLog = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & _
                                Application.PathSeparator & "Erreurs.log"
    
    'Ouverture de log des erreurs
    Dim canalErreurLog As Integer
    canalErreurLog = FreeFile
    Open fullPathErreurLog For Append As #canalErreurLog
    
    'Écrire dans le fichier
    Print #canalErreurLog, ligneLog
    
    Close #canalErreurLog

    On Error GoTo 0
    
End Sub

'Public Sub InitialiserMenuPrincipal(Optional source As String = "Appel inconnu") '2025-11-10 @ 08:47
'
'    Dim startTime As Double: startTime = Timer
'    Call modDev_Utils.EnregistrerLogApplication("modAppli:InitialiserMenuPrincipal — Source : " & source, vbNullString, 0)
'
'    'Vérification d’ouverture silencieuse
'    Call modTraceSession.VerifierOuvertureSilencieuse
'
'    Call CacherToutesFeuillesSaufMenu
'
'    'Mise à jour des formes selon l’utilisateur
'    Call modMenu.CacherFormesEnFonctionUtilisateur(Fn_UtilisateurWindows)
'
'    'Protection de la feuille menu
'    With wshMenu
'        .Protect UserInterfaceOnly:=True
'        .EnableSelection = xlUnlockedCells
'    End With
'
'    'Positionnement visuel
'    wshMenu.Activate
'
'    'Journalisation de la durée
'    Call modDev_Utils.EnregistrerLogApplication("modAppli:InitialiserMenuPrincipal terminé", vbNullString, startTime)
'
'End Sub
'
Public Sub InitialiserFeuilleMenu(ws As Worksheet, contexte As String, etiquette As String) '2025-11-11 @ 10:01
    
    Call modSessionVerrou.VerrouillerSiSessionInvalide("Activation Feuille '" & ws.Name & "'")
    
    If Not gSessionInitialisee Then
        MsgBox "L'application n'a pas été initialisée correctement.", vbCritical, etiquette
        Call FermerApplication("Activate:" & ws.Name, True)
        Exit Sub
    End If
    
    With ws
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Call modTraceSession.VerifierOuvertureSilencieuse
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli:InitialiserFeuilleMenu", etiquette, -1)
    
End Sub

