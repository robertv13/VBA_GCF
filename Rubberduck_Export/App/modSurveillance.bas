Attribute VB_Name = "modSurveillance"
Option Explicit

Public gDerniereActiviteNomFeuille As String
Public gDerniereActiviteAdresseSelection As String
Public gDerniereInteractionTimer As Double

Sub LancerSurveillance() '2025-11-06 @ 15:57

    gHeureProchaineVerification = Now + TimeSerial(0, gFREQUENCE_VERIFICATION_INACTIVITE, 0)
    Application.OnTime gHeureProchaineVerification, "modSurveillance.VerifierActivite" 'Aux 5 minutes
    
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
        Application.StatusBar = "Prochaine vérification d'activité prévue à " & gHeureProchaineVerification
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
        Debug.Print Now() & " Ça bouge - La feuille a changé"
        ActiviteDetectee = True
    End If
    
    If Selection.Address <> gDerniereActiviteAdresseSelection Then
        Debug.Print Now() & " Ça bouge - La sélection a changé"
        ActiviteDetectee = True
    End If
    
    If Timer - gDerniereInteractionTimer < (gMAXIMUM_INACTIVITE * 60) Then
        Debug.Print Now() & " [ActiviteDetectee] Il y a eu au minimum UNE activité dans les " & gMAXIMUM_INACTIVITE & " dernières minutes."
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

    Debug.Print Now() & " Aucune activité, on demande à l'utilisateur de fermer ou de garder l'application active ?"
    Dim choix As VbMsgBoxResult
    choix = MsgBox("Aucune activité détectée depuis au moins " & gMAXIMUM_INACTIVITE & " minutes." & _
                    vbCrLf & vbCrLf & "Souhaitez-vous fermer l'application ?", _
                    vbYesNoCancel + vbCritical, _
                    "APP GCF - Aucune activité - Fermeture automatique")

    Select Case choix
        Case vbYes
            Call modSurveillance.FermerApplicationConfirme
        Case vbNo
            Call LancerSurveillance 'Relance après délai
        Case vbCancel
            gDerniereInteractionTimer = Timer 'Reset
    End Select
    
End Sub

Sub FermerApplicationConfirme()  '2025-11-06 @ 15:57

    Call modMenu.FermerApplication("Fermeture confirmée par l'utilisateur", False)
    
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
            Case "ListBox"
                Set wrapper = New clsWrapperListBox
                Set wrapper.lb = ctrl
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
    ligne = Now() & " | " & source & " | Feuille: " & Fn_NomFeuilleActive()
    
    Call EcrireDansLog(ligne)
    
End Sub

Public Sub EcrireDansLog(texte As String) '2025-11-07 @ 04:36

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

Public Function Fn_NomFeuilleActive() As String '2025-11-07 @ 04:38

    On Error Resume Next
    Fn_NomFeuilleActive = ActiveSheet.Name
    
End Function

