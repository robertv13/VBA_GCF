Attribute VB_Name = "modSurveillance"
Option Explicit

Public gDerniereActiviteNomFeuille As String
Public gDerniereActiviteAdresseSelection As String
Public gDerniereInteractionTimer As Double
Public gTimerFermetureActif As Boolean
Public gProchainTick As Date                    'Prochain compte à rebours
Public gProchainRafraichir As Date                    'Prochain compte à rebours
Public fermetureAuto As clsFermetureAuto

Sub LancerSurveillance() '2025-11-06 @ 15:57

    gHeureProchaineVerification = Now + TimeSerial(0, gFREQUENCE_VERIFICATION_INACTIVITE, 0)
    Application.OnTime gHeureProchaineVerification, "modSurveillance.VerifierActivite" 'Aux 5 minutes
    Debug.Print Now() & " La prochaine vérification se fera à " & gHeureProchaineVerification
    
End Sub

Sub VerifierActivite() '2025-11-06 @ 15:57

    If Hour(Now) < gHEURE_DEBUT_SURVEILLANCE Then Exit Sub

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modSurveillance:VerifierActivite", vbNullString, 0)
    
    Debug.Print Now() & " VerifierActivite"
    If Not ActiviteDetectee() Then
        Call ProposerFermeture
    Else
        Dim prochaineVerification As Date
        prochaineVerification = gHeureProchaineVerification + TimeSerial(0, gFREQUENCE_VERIFICATION_INACTIVITE, 0)
        Application.StatusBar = "Vérification d'inactivité faite - Prochaine vérification d'inactivité prévue à " & Right(prochaineVerification, 8)
        Call LancerSurveillance 'Relance la surveillance
    End If
    
    Call modDev_Utils.EnregistrerLogApplication("modSurveillance:VerifierActivite", vbNullString, _
                                                        startTime)

End Sub

Function ActiviteDetectee() As Boolean '2025-11-06 @ 15:57

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modSurveillance:ActiviteDetectee", vbNullString, 0)
    
    ActiviteDetectee = False
    
    'Vérification de contexte : l'utilisateur est-il dans le classeur de l'application ?
    If Not ActiveWorkbook Is ThisWorkbook Then
        Debug.Print Now() & " L'utilisateur est dans un autre classeur"
        GoTo FIN_VERIFICATION
    End If
    
    If ActiveSheet.Name <> gDerniereActiviteNomFeuille Then
        Debug.Print Now() & " Ça bouge - Une feuille a changé"
        ActiviteDetectee = True
    End If
    
    If Selection.Address <> gDerniereActiviteAdresseSelection Then
        Debug.Print Now() & " Ça bouge - Une sélection a changé"
        ActiviteDetectee = True
    End If
    
    If Timer - gDerniereInteractionTimer < (gMAXIMUM_INACTIVITE * 60) Then
        Debug.Print Now() & " Il y a eu au minimum UNE activité dans les " & gMAXIMUM_INACTIVITE & " dernières minutes."
        ActiviteDetectee = True
    End If
    
FIN_VERIFICATION:

    If ActiviteDetectee = False Then
        Debug.Print Now() & " A U C U N E   A C T I V I T É !"
    Else
        gDerniereActiviteNomFeuille = ActiveSheet.Name
        gDerniereActiviteAdresseSelection = Selection.Address
    End If
    
    Call modDev_Utils.EnregistrerLogApplication("modSurveillance:ActiviteDetectee", vbNullString, _
                                                        startTime)

End Function

Sub ProposerFermeture() '2025-11-06 @ 15:57

    Set fermetureAuto = New clsFermetureAuto
    Call AfficherFormulaireUrgent(gMAXIMUM_INACTIVITE)
    
End Sub

Public Sub InitialiserSurveillanceForm(frm As Object, ByRef wrappers As Collection) '2025-11-06 @ 16:21

    Set wrappers = New Collection
    Dim ctrl As MSForms.Control
    Dim wrapper As Object

    For Each ctrl In frm.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                Set wrapper = New clsWrapperTextBox
                Set wrapper.tb = ctrl
            Case "ComboBox"
                Set wrapper = New clsWrapperComboBox
                Set wrapper.cb = ctrl
            Case "CommandButton"
                Set wrapper = New clsWrapperCommandButton
                Set wrapper.btn = ctrl
'            Case "ListBox"
'                Set wrapper = New clsWrapperListBox
'                Set wrapper.lb = ctrl
            Case "CheckBox"
                Set wrapper = New clsWrapperCheckBox
                Set wrapper.chk = ctrl
            Case "OptionButton"
                Set wrapper = New clsWrapperOptionButton
                Set wrapper.opt = ctrl
            Case Else
                Set wrapper = Nothing
        End Select
        If Not wrapper Is Nothing Then wrappers.Add wrapper
    Next ctrl
    
End Sub

Public Sub EnregistrerActivite(source As String) '2025-11-06 @ 17:36

    gDerniereInteractionTimer = Timer
    Dim ligne As String
    ligne = Format$(Now(), "yyyy-mm-dd hh:nn:ss") & " | " & source & " | Feuille: " & Fn_NomFeuilleActive()
    
    Call EnregistrerActiviteDurantSurveillance(ligne)
    
End Sub

Public Sub EnregistrerActiviteDurantSurveillance(texte As String) '2025-11-07 @ 04:36

    If Hour(Now) < gHEURE_DEBUT_SURVEILLANCE Then Exit Sub
    
    Dim chemin As String
    chemin = ThisWorkbook.path & "\journal_activite.txt"
    
    Dim fso As Object, fichier As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.fileExists(chemin) Then
        Set fichier = fso.OpenTextFile(chemin, 8) 'Append
    Else
        Set fichier = fso.CreateTextFile(chemin, True)
    End If
    
    fichier.WriteLine texte
    fichier.Close
    
End Sub

Public Sub InitierFermetureSilencieuse(Optional minutesInactives As Double = 0) '2025-11-08 @ 06:20

    'Affiche le formulaire avec délai de grâce
    If ufConfirmationFermeture.Visible = False Then
        Call ufConfirmationFermeture.AfficherMessageFermetureAPP(minutesInactives)
        gProchainTick = Now + TimeSerial(0, 0, 1)
        Application.OnTime gProchainTick, "modSurveillance.SurveillerFermetureAuto"
    End If
    
End Sub

Public Sub SurveillerFermetureAuto()

    If fermetureAuto Is Nothing Then Exit Sub
    fermetureAuto.Rafraichir
    
End Sub

Public Function Fn_NomFeuilleActive() As String '2025-11-07 @ 04:38

    On Error Resume Next
    Fn_NomFeuilleActive = ActiveSheet.Name
    
End Function

Public Sub AfficherFormulaireUrgent(Optional minutesInactives As Double = 0) '2025-11-08 @ 07:26

    On Error Resume Next

    '1. Ramener Excel au premier plan
    AppActivate Application.Caption

    '2. S'assurer que la fenêtre Excel est visible
    Application.Visible = True
'    Application.WindowState = xlNormal

    '3. Centrer le formulaire dans Excel
    With ufConfirmationFermeture
        .StartUpPosition = 0
        .Left = Application.Left + (Application.Width - .Width) / 2
        .Top = Application.Top + (Application.Height - .Height) / 2 + 100

    '4. Préparer le message
        .lblMessage.Caption = "Aucune activité détectée depuis " & Format$(minutesInactives, "0") & " minutes..." & vbCrLf & vbCrLf & _
                              "Souhaitez-vous garder l’application ouverte quand même ?"
        .lblTimer.Caption = vbNullString

    '5. Afficher le formulaire
        .show vbModeless
    End With

    '6. Lancer le décompte visuel
    Set fermetureAuto = New clsFermetureAuto
    Call fermetureAuto.FermetureDansXSecondes(gDELAI_GRACE_SECONDES)
    
End Sub

Public Sub PlanifierTick()

    gProchainTick = Now + TimeSerial(0, 0, 1)
    Application.OnTime gProchainTick, "modSurveillance.SurveillerFermetureAuto"
    
End Sub

Public Sub AnnulerTick()

    On Error Resume Next
    Application.OnTime gProchainTick, "modSurveillance.SurveillerFermetureAuto", , False
    
End Sub

Sub TesterFermeture() '2025-11-08 @ 06:37

    Set fermetureAuto = New clsFermetureAuto
    fermetureAuto.SimulerFermetureDansXSecondes 20 'Test avec 15 secondes
    
End Sub

