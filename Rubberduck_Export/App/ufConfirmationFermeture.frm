VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufConfirmationFermeture 
   Caption         =   "Confirmation de fermeture de l'application"
   ClientLeft      =   120
   ClientTop       =   465
   OleObjectBlob   =   "ufConfirmationFermeture.frx":0000
End
Attribute VB_Name = "ufConfirmationFermeture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tFermeture As Date
Public tProchainTick As Date
Dim gClignoteEtat As Boolean

Private Sub UserForm_Initialize()

    Me.BackColor = RGB(245, 245, 245) 'Gris clair
    
    Me.lblMessage.BackColor = RGB(255, 255, 255)
    Me.lblMessage.ForeColor = RGB(0, 0, 0)
    
    Me.cmdGarderOuverte.BackColor = RGB(210, 255, 210)
    Me.cmdGarderOuverte.ForeColor = RGB(0, 100, 0)
    
    Me.cmdFermerMaintenant.BackColor = RGB(210, 255, 210)
    Me.cmdFermerMaintenant.ForeColor = RGB(160, 0, 0)
    
End Sub

Private Sub cmdGarderOuverte_Click() '2025-07-01 @ 17:13
    
    Debug.Print "[cmdGarderOuverte_Click] Utilisateur a cliqué sur 'Garder l'application ouverte' à : " & Format(Now, "hh:mm:ss")
    
    On Error Resume Next
    
    'Annule la fermeture automatique planifiée
    Application.OnTime gFermeturePlanifiee, "FermerApplicationAucuneActivite", , False
    
    'Annule le clignotement du timer (si encore actif)
    Application.OnTime tProchainTick, "RelancerTimer", , False
        
    On Error GoTo 0
    
    'Réinitialise le timestamp d'activité
    gDerniereActivite = Now
    
    'Nettoie le formulaire (optionnel mais propre)
    lblMessage.Caption = vbNullString
    lblTimer.Caption = vbNullString
    gClignoteEtat = False
    
    'Ferme le UserForm
    Me.Hide
    
End Sub

Private Sub cmdFermerMaintenant_Click() '2025-07-01 @ 15:46

    Debug.Print "[cmdFermerMaintenant_Click] Utilisateur a cliqué sur 'Fermer maintenant' à : " & Format(Now, "hh:mm:ss")
    
    Me.Hide
    Call FermerApplicationNormalement(GetNomUtilisateur())
    
End Sub

Public Sub afficherMessage(Optional minutesInactives As Double = 0) '2025-07-01 @ 15:56

    Dim msg As String
    msg = "Aucune activité détectée depuis " & Format$(minutesInactives, "0") & " minutes..." & vbCrLf & vbCrLf
    msg = msg & "Souhaitez-vous garder l’application ouverte quand même ?"

    lblMessage.Caption = msg
    
    tFermeture = Now + TimeSerial(0, 0, gDELAI_GRACE_SECONDES)
    gFermeturePlanifiee = tFermeture
    Debug.Print "[AfficherMessage] gFermeturePlanifiee synchronisé à : " & Format(gFermeturePlanifiee, "hh:mm:ss")
    lblTimer.Caption = vbNullString
    Debug.Print "[AfficherMessage] Affichage du formulaire de confirmation à : " & Format(Now, "hh:mm:ss")
    Debug.Print "[AfficherMessage] Fermeture prévue à (tFermeture) : " & Format(tFermeture, "hh:mm:ss")
    Call ufConfirmationFermeture.RafraichirTimer
    
    Me.StartUpPosition = 1
    Me.show vbModeless
    
End Sub

Public Sub RafraichirTimer() '2025-07-02 @ 06:56

    If tFermeture = 0 Then
        Debug.Print "RafraichirTimer déclenché alors que tFermeture = 0 — arrêt immédiat"
        Exit Sub
    End If
    
    Dim delta As Double
    delta = DateDiff("s", Now, tFermeture)
    
    'Journal : moment d’exécution et delta
    Debug.Print "[RafraichirTimer] RafraichirTimer à " & Format(Now, "hh:mm:ss") & _
                " | tFermeture : " & Format(tFermeture, "hh:mm:ss") & _
                " | Secondes restantes : " & delta
                
    If delta <= 0 Then
        lblTimer.Caption = "Temps écoulé — fermeture en cours..."
        lblTimer.ForeColor = RGB(120, 0, 0)
        'Journal : fin du countdown
        Debug.Print "Temps écoulé — arrêt du timer visuel"
        Exit Sub
    End If

    'Mise à jour du texte
    lblTimer.Caption = "Fermeture automatique dans " & _
        Format$(delta \ 60, "00") & ":" & Format$(delta Mod 60, "00")
    
    'Clignotement si < 60 s
    If delta <= 30 Then
        gClignoteEtat = Not gClignoteEtat
        If gClignoteEtat Then
            lblTimer.ForeColor = RGB(200, 0, 0) 'Rouge vif
        Else
            lblTimer.ForeColor = RGB(255, 255, 255) 'Invisible (blanc sur fond clair)
        End If
        'Journal : clignotement actif
'        Debug.Print "[ufConfirmationFermeture:RafraichirTimer] Clignotement actif — gClignoteEtat = " & gClignoteEtat
    Else
        lblTimer.ForeColor = RGB(0, 0, 128) 'Bleu foncé normal
    End If

    'Replanification dans 1 s
    tProchainTick = Now + TimeSerial(0, 0, 1)
    'Journal : programmation du prochain appel
    Debug.Print "[RafraichirTimer] Prochain appel prévu à : " & Format(tProchainTick, "hh:mm:ss")
    Application.OnTime tProchainTick, "RelancerTimer"
    
End Sub

Public Function ProchainTick() As Date '2025-07-02 @ 08:19

    ProchainTick = tProchainTick
    
End Function


