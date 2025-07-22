Attribute VB_Name = "modAppli"
Option Explicit

Private Sub Auto_Open() '2024-12-28 @ 11:09

    'Chemin du dossier contenant les fichiers PROD - 2025-07-02 @ 14:24
    Dim cheminSourcePROD As String
    cheminSourcePROD = "P:\Administration\APP\GCF\DataFiles\"
    
    If Fn_Get_Windows_Username <> "RobertMV" And Fn_Get_Windows_Username <> "Robertmv" Then
        If Dir(cheminSourcePROD & "\GCF_BD_MASTER.lock") <> vbNullString Then
            MsgBox "Cette application est actuellement en maintenance." & vbNewLine & vbNewLine & _
                   "Le fichier principal est verrouillé par le développeur." & vbNewLine & vbNewLine & _
                   "Veuillez ressayer dans 5 à 10 minutes SVP", _
                   vbCritical, _
                   "L'application APP n'est pas disponible"
            Call modMenu.FermerApplicationNormalement(modFunctions.GetNomUtilisateur())
        End If
    End If
    
    gDerniereActivite = Now

    'Mise en placedu mécanisme pour sortir automatiquement de l'application, s'il n'y a pas d'activité
    gProchaineVerification = Now + TimeSerial(0, gFREQUENCE_VERIFICATION_INACTIVITE, 0)
    Application.OnTime gProchaineVerification, "VerifierDerniereActivite"
    Application.EnableEvents = False
    wsdADMIN.Range("B3").Value = gFREQUENCE_VERIFICATION_INACTIVITE
    wsdADMIN.Range("B4").Value = gMAXIMUM_MINUTES_INACTIVITE
    Application.EnableEvents = True

    Call DemarrerApplication
    
End Sub

Sub DemarrerApplication() '2025-07-11 @ 15:16

    Dim rootPath As String
    rootPath = ObtenirRepertoireBase

    Application.EnableEvents = False
    wsdADMIN.Range("F5").Value = rootPath
    Application.EnableEvents = True
   
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("----- DÉBUT D'UNE NOUVELLE SESSION (modAppli:DemarrerApplication) -----", vbNullString, 0)
    
    'Quel est l'utilisateur Windows ?
    gUtilisateurWindows = modFunctions.GetNomUtilisateur()
    
    On Error GoTo ErrorHandler
    
    If Application.EnableEvents = False Then Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.StatusBar = "Vérification de l'accès au répertoire principal"
    If Fn_Check_Server_Access(rootPath) = False Then
        MsgBox "Le répertoire principal '" & rootPath & "' n'est pas accessible." & vbNewLine & vbNewLine & _
               "Veuillez vérifier votre connexion au serveur SVP", vbCritical, rootPath
        Exit Sub
    End If
    Application.StatusBar = False

    Call CreerFichierUtilisateurActif(gUtilisateurWindows)
    Call FixerFormatDateUtilisateur(gUtilisateurWindows)
    Call CreerSauvegardeMaster
    Call EcrireInformationsConfigAuMenu(gUtilisateurWindows)
    wshMenu.Range("A1").Value = wsdADMIN.Range("NomEntreprise").Value
    Call modMenu.HideDevShapesBasedOnUsername(gUtilisateurWindows)
    
    'Protection de la feuille wshMenu
    With wshMenu
        .Protect userInterfaceOnly:=True
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

    If gUtilisateurWindows = "RobertMV" Or gUtilisateurWindows = "robertmv" Then
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

Function ObtenirRepertoireBase() As String '2025-03-03 @ 20:28
   
    DoEvents
    
    If modFunctions.GetNomUtilisateur() = "RobertMV" Or modFunctions.GetNomUtilisateur() = "robertmv" Then
        ObtenirRepertoireBase = "C:\VBA\GC_FISCALITÉ"
    Else
        ObtenirRepertoireBase = "P:\Administration\APP\GCF"
    End If

End Function

Sub CreerFichierUtilisateurActif(ByVal userName As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli:CreerFichierUtilisateurActif", vbNullString, 0)
    
    Dim traceFilePath As String
    traceFilePath = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & "Actif_" & userName & ".txt"
    
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
    
    Select Case user
        Case "GuillaumeCharron", "Guillaume", "gchar"
            userDateFormat = "dd/mm/yy"
        Case "vgervais", "Vlad_Portable", "User", "Oli_Portable"
            userDateFormat = "dd/mm/yyyy"
        Case "Annie"
            userDateFormat = "yyyy/mm/dd"
        Case "RobertMV", "robertmv"
            userDateFormat = "dd/mm/yyyy"
        Case Else
            userDateFormat = "dd/mm/yyyy"
    End Select

    wsdADMIN.Range("B1").Value = userDateFormat
    
    Call modDev_Utils.EnregistrerLogApplication("modAppli:FixerFormatDateUtilisateur", vbNullString, startTime)
    
End Sub

Sub CreerSauvegardeMaster()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli:CreerSauvegardeMaster", vbNullString, 0)
    
    On Error GoTo MASTER_NOT_AVAILABLE
    
    'Chemin source (fichier principal) et destination (sauvegarde)
    Dim masterFilePath As String
    masterFilePath = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & "GCF_BD_MASTER.xlsx"
    
    Dim backupFilePath As String
    backupFilePath = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    
    'Créer directement une copie du fichier sans ouvrir Excel
    FileCopy masterFilePath, backupFilePath

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

Sub EcrireInformationsConfigAuMenu(ByVal user As String)
    
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
    environnement = wsdADMIN.Range("F5").Value

    valeurs = Array( _
        "Heure - " & Format$(Now(), formatDate & " hh:mm:ss"), _
        "Version - " & ThisWorkbook.Name, _
        "Utilisateur - " & user, _
        "Environnement - " & environnement, _
        "Format de la date - " & formatDate)

    ' Ecriture en une seule opération
    wshMenu.Range("A30:A34").Value = Application.WorksheetFunction.Transpose(valeurs)

    With wshMenu
        .Protect userInterfaceOnly:=True
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
    Call EnregistrerActiviteAuLog("[modAppli:EnregistrerActivite] " & msg) '2025-07-03 @ 10:31
    
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
        If gMODE_DEBUG Then Debug.Print "[modAppli:VerifierDerniereActivite] gDerniereActivite non initialisée"
        Exit Sub
    End If

    'Calcul du temps d'inactivité en minutes
    Dim minutesInactives As Double
    minutesInactives = Round(MinutesDepuisDerniereActivite(), 1)

    If gMODE_DEBUG Then Debug.Print "[modAppli:VerifierDerniereActivite] Vérification @ " & minutesInactives & " min. - " & _
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
        If gMODE_DEBUG Then Debug.Print "[modAppli:VerifierDerniereActivite] Inactivité trop longue (" & Format$(minutesInactives, "0") & " minutes) — fermeture de l'application"
        gFermeturePlanifiee = GetProchaineFermeture()
        If gMODE_DEBUG Then Debug.Print "[modAppli:VerifierDerniereActivite] Avant l'ajout de " & gDELAI_GRACE_SECONDES & " secondes, gFermeture = " & Format(gFermeturePlanifiee, "; hh: mm: ss "); vbNullString
        gFermeturePlanifiee = Now + TimeSerial(0, 0, gDELAI_GRACE_SECONDES)
        If gMODE_DEBUG Then Debug.Print "[modAppli:VerifierDerniereActivite] Après l'ajout de " & gDELAI_GRACE_SECONDES & " secondes, gFermeture = " & Format(gFermeturePlanifiee, "hh:mm:ss")
        
        On Error Resume Next
        If gMODE_DEBUG Then Debug.Print "[modAppli:VerifierDerniereActivite] OnTime prévu pour : " & Format(gFermeturePlanifiee, "hh:mm:ss")
        Application.OnTime gFermeturePlanifiee, "FermerApplicationAucuneActivite"
        On Error GoTo 0
        
        Unload ufConfirmationFermeture '2025-07-02 @ 07:54
        
        Call ufConfirmationFermeture.afficherMessage(minutesInactives)
        Exit Sub
    End If

    'Replanification
    Call PlanifierVerificationDerniereActivite
    Exit Sub

GestionErreur:
    Debug.Print "[modAppli:VerifierDerniereActivite] Erreur dans VerifierDerniereActivite : " & Err.Number & " - " & Err.description

End Sub

Public Sub PlanifierVerificationDerniereActivite() '2025-07-01 @ 13:53
    
    On Error Resume Next
    Application.OnTime gProchaineVerification, "VerifierDerniereActivite", , False
    On Error GoTo 0
    
    gProchaineVerification = Now + TimeSerial(0, gFREQUENCE_VERIFICATION_INACTIVITE, 0)
    Application.OnTime gProchaineVerification, "VerifierDerniereActivite"
    If gMODE_DEBUG Then Debug.Print "[modAppli:PlanifierVerificationDerniereActivite] Prochaine vérification à " & Format(gProchaineVerification, "hh:mm:ss")
    
End Sub

Public Sub FermerApplicationAucuneActivite() '2025-07-02 @ 06:19

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modAppli:FermerApplicationAucuneActivite", vbNullString, 0)
    
    'Ajoute un log pour vérification
    If gMODE_DEBUG Then Debug.Print "[modAppli:FermerApplicationAucuneActivite] Fermeture automatique déclenchée à : " & Format(Now, "hh:mm:ss")

    'Appel direct de la procédure de fermeture
    Call FermerApplicationNormalement(modFunctions.GetNomUtilisateur())
    
End Sub

Public Sub RelancerTimer() '2025-07-02 @ 06:43

    If gMODE_DEBUG Then Debug.Print "[modAppli:RelancerTimer] Appel de 'ufConfirmationFermeture.RafraichirTimer'"
    ufConfirmationFermeture.RafraichirTimer
    
End Sub

Public Sub RelancerSurveillance() '2025-07-02 @ 07:41

    If gMODE_DEBUG Then Debug.Print "[modAppli:RelancerSurveillance] *** Surveillance relancée manuellement à " & Format(Now, "hh:mm:ss")
    
    On Error Resume Next
    If gFermeturePlanifiee = 0 Then
        If gMODE_DEBUG Then Debug.Print "[modAppli:RelancerSurveillance] gFermeturePlanifiee est nul — aucun OnTime à annuler"
    End If

    Application.OnTime gFermeturePlanifiee, "FermerApplicationAucuneActivite", , False
    Application.OnTime ufConfirmationFermeture.ProchainTick, "RelancerTimer", , False
    On Error GoTo 0

    gDerniereActivite = Now
    Call VerifierDerniereActivite

End Sub

Public Sub EnregistrerActiviteAuLog(ByVal message As String) '2025-07-03 @ 10:29

    'Ne rien faire avant l'heure de début de la surveillance
    If TimeValue(Now) < TimeSerial(gHEURE_DEBUT_SURVEILLANCE, 0, 0) Then
        Exit Sub
    End If

    Dim cheminLog As String
    cheminLog = wsdADMIN.Range("F5").Value & gDATA_PATH & "\ActiviteDurantSurveillance.txt"

    Dim fileNum As Integer
    fileNum = FreeFile

    Open cheminLog For Append As #fileNum
    Print #fileNum, "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] [" & Fn_Get_Windows_Username & "] " & message
    Close #fileNum
    
End Sub
