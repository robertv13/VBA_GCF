VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFermetureAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Classe : clsFermetureAuto

Private heureLimite As Date
Private timerActif As Boolean
Private minutesInactives As Double

Public Sub Initialiser(Optional minutes As Double = 0)

    minutesInactives = minutes
    heureLimite = Now + TimeSerial(0, 0, gDELAI_GRACE_SECONDES)
    timerActif = True

    With ufConfirmationFermeture
        .lblMessage.Caption = "Aucune activité détectée depuis au moins " & Format$(minutes, "0") & " minutes..." & vbCrLf & vbCrLf & _
                              "Souhaitez-vous garder l’application ouverte quand même ?"
        .lblTimer.Caption = vbNullString
        .StartUpPosition = 1
        .show vbModeless
    End With

    Debug.Print Now() & " [clsFermetureAuto] Fermeture prévue à : " & Format(heureLimite, "hh:mm:ss")
    gProchainTick = Now + TimeSerial(0, 0, 1)
    Application.OnTime gProchainTick, "modSurveillance.SurveillerFermetureAuto"
    
End Sub

Public Sub Rafraichir()

    If Not timerActif Then Exit Sub

    Dim secondesRestantes As Long
    secondesRestantes = DateDiff("s", Now, heureLimite)

    If secondesRestantes <= 0 Then
        timerActif = False
        Unload ufConfirmationFermeture
        Call modMenu.FermerApplication("Fermeture forcée après TimeOut", False)
    Else
        ufConfirmationFermeture.lblTimer.Caption = "Fermeture dans " & _
                                    Format$(secondesRestantes \ 60, "00") & ":" & _
                                    Format$(secondesRestantes Mod 60, "00")
        Call modSurveillance.PlanifierTick
    End If

End Sub

Public Sub Annuler()

    timerActif = False
    Call modSurveillance.AnnulerTick

End Sub

Public Sub FermetureDansXSecondes(delaiSecondes As Long)

    minutesInactives = 0
    heureLimite = Now + TimeSerial(0, 0, delaiSecondes)
    timerActif = True

    With ufConfirmationFermeture
        .lblMessage.Caption = "Fermeture dans " & delaiSecondes & " secondes..." & vbCrLf & vbCrLf & _
                              "Souhaitez-vous garder l’application ouverte quand même ?"
        .lblTimer.Caption = vbNullString
        .StartUpPosition = 0
        .show vbModeless
    End With

    Debug.Print Now() & " [SimulerFermeture] Fermeture simulée à : " & Format(heureLimite, "hh:mm:ss")
    gProchainTick = Now + TimeSerial(0, 0, 1)
    Application.OnTime gProchainTick, "modSurveillance.SurveillerFermetureAuto"
    
End Sub

Public Sub SimulerFermetureDansXSecondes(delaiSecondes As Long)

    minutesInactives = 0
    heureLimite = Now + TimeSerial(0, 0, delaiSecondes)
    timerActif = True

    With ufConfirmationFermeture
        .lblMessage.Caption = "SIMULATION : Fermeture dans " & delaiSecondes & " secondes..." & vbCrLf & vbCrLf & _
                              "Souhaitez-vous garder l’application ouverte quand même ?"
        .lblTimer.Caption = vbNullString
        .StartUpPosition = 0
        .show vbModeless
    End With

    Debug.Print Now() & " [SimulerFermeture] Fermeture simulée à : " & Format(heureLimite, "hh:mm:ss")
    gProchainTick = Now + TimeSerial(0, 0, 1)
    Application.OnTime gProchainTick, "modSurveillance.SurveillerFermetureAuto"
    
End Sub

