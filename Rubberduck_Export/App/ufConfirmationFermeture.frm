VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufConfirmationFermeture 
   Caption         =   "Confirmation AVANT la fermeture de l'application"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7350
   OleObjectBlob   =   "ufConfirmationFermeture.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufConfirmationFermeture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Me.BackColor = RGB(245, 245, 245) 'Gris clair
    
    Me.lblMessage.BackColor = RGB(255, 255, 255)
    Me.lblMessage.ForeColor = RGB(0, 0, 0)
    
    Me.btnGarderOuverte.BackColor = RGB(210, 255, 210)
    Me.btnGarderOuverte.ForeColor = RGB(0, 100, 0)
    
    Me.btnFermerMaintenant.BackColor = RGB(210, 255, 210)
    Me.btnFermerMaintenant.ForeColor = RGB(160, 0, 0)
    
    'Positionnement manuel
    Me.StartUpPosition = 0

End Sub

Private Sub btnFermerMaintenant_Click() '2025-11-08 @ 06:11

    Debug.Print Now() & " [btnFermerMaintenant_Click] Utilisateur a cliqué sur 'Fermer maintenant' à : " & Format(Now, "hh:mm:ss")
    
    fermetureAuto.Annuler
    
    Unload Me
    Call modSurveillance.FermerApplicationConfirme
    gTimerFermetureActif = False
    Application.StatusBar = False

    Call modMenu.FermerApplication("Application inactive - Fermeture souhaitée", False)
    
End Sub

Private Sub btnGarderOuverte_Click() '2025-11-08 @ 06:16
    
    Debug.Print Now() & " [btnGarderOuverte_Click] Utilisateur a cliqué sur 'Garder l'application ouverte' à : " & Format(Now, "hh:mm:ss")
    
    fermetureAuto.Annuler
    
    Unload Me
    Call modSurveillance.LancerSurveillance
    gTimerFermetureActif = False
    Application.StatusBar = False
    
End Sub

Public Sub AfficherMessageFermetureAPP(Optional minutesInactives As Double = 0) '2025-11-08 @ 05:58

    Dim msg As String
    msg = "Aucune activité de détectée depuis " & Format$(minutesInactives, "0") & " minutes..." & vbCrLf & vbCrLf
    msg = msg & "Souhaitez-vous garder l’application ouverte quand même ?"

    lblMessage.Caption = msg
    
    gHeurePrevueFermetureAutomatique = Now + TimeSerial(0, 0, gDELAI_GRACE_SECONDES)
    lblTimer.Caption = vbNullString
    
    Call ufConfirmationFermeture.RafraichirTimer

    Me.StartUpPosition = 1
    Me.show vbModeless

    Call DémarrerTimerVisuel

End Sub

Public Sub RafraichirTimer() '2025-07-02 @ 06:56

    If Not gTimerFermetureActif Then Exit Sub

    Dim secondesRestantes As Long
    secondesRestantes = DateDiff("s", Now, gHeurePrevueFermetureAutomatique)

    If secondesRestantes <= 0 Then
        lblTimer.Caption = "Fermeture imminente..."
        gTimerFermetureActif = False
        Unload Me
        Call modSurveillance.FermerApplicationConfirme
    Else
        lblTimer.Caption = "Fermeture dans " & Format$(secondesRestantes \ 60, "00") & ":" & _
                                                    Format$(secondesRestantes Mod 60, "00") & "..."
        gProchainRafraichir = Now + TimeSerial(0, 0, 1)
        Application.OnTime gProchainRafraichir, "ufConfirmationFermeture.RafraichirTimer"
    End If
    
End Sub

Public Function ProchainTick() As Date '2025-07-02 @ 08:19

    ProchainTick = gProchainTick
    
End Function

Public Sub DémarrerTimerVisuel() '2025-11-08 @ 06:26

    gTimerFermetureActif = True
    Call RafraichirTimer
    
End Sub

