Attribute VB_Name = "modTraceSession"
Option Explicit '2025-11-10 @ 08:35 Tenter de solutionner crash / redémarrages intempestifs

Public gSessionInitialisee As Boolean

Public Const NOM_FICHIER_OUVERTURE_NORMALE As String = "OuvertureNormale"
Public Const NOM_FICHIER_SESSION_ACTIVE As String = "SessionActive"

Public Sub CreerTraceOuverture()

    On Error Resume Next
        Dim chemin As String: chemin = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
        Dim f As Integer: f = FreeFile
        Open chemin & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & _
                Fn_NomFichierControleSession(NOM_FICHIER_OUVERTURE_NORMALE) For Output As #f
        Print #f, "Ouverture normale à " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " par " & _
                                                                                    Fn_UtilisateurWindows
        Close #f
    On Error GoTo 0

End Sub

Public Sub SupprimerTraceOuverture()

    On Error Resume Next
        Dim chemin As String: chemin = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
        Kill chemin & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & _
                Fn_NomFichierControleSession(NOM_FICHIER_OUVERTURE_NORMALE)
    On Error GoTo 0
    
End Sub

Public Sub VerifierOuvertureSilencieuse()

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modTraceSession:VerifierOuvertureSilencieuse", _
                                                                                        vbNullString, 0)
    
'    On Error Resume Next
        Dim chemin As String: chemin = Fn_RepertoireBaseApplication(gUtilisateurWindows)
        If Dir(chemin & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & _
                    Fn_NomFichierControleSession(NOM_FICHIER_OUVERTURE_NORMALE)) <> vbNullString Then
            If Timer - FileDateTime(chemin & Application.PathSeparator & gDATA_PATH & _
                                    Application.PathSeparator & _
                                    Fn_NomFichierControleSession(NOM_FICHIER_OUVERTURE_NORMALE)) > 5 Then
                'Le fichier existe mais Workbook_Open n’a pas été exécuté
                Call AppelLogSessionAnormale("Ouverture silencieuse détectée — Workbook_Open non exécuté")
            End If
        End If
'    On Error GoTo 0
    
    Call modDev_Utils.EnregistrerLogApplication("modTraceSession:VerifierOuvertureSilencieuse", _
                                                                                vbNullString, startTime)

End Sub

Private Sub AppelLogSessionAnormale(msg As String)

    On Error Resume Next
        Dim chemin As String: chemin = Fn_RepertoireBaseApplication(gUtilisateurWindows)
        Dim f As Integer: f = FreeFile
        Open chemin & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & _
                                Fn_NomFichierControleSession(NOM_FICHIER_SESSION_ACTIVE) For Append As #f
        Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & gUtilisateurWindows & " | " & msg
        Close #f
    On Error GoTo 0
    
End Sub

