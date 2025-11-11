Attribute VB_Name = "modTraceSession"
Option Explicit '2025-11-10 @ 08:35 Tenter de solutionner crash / redémarrages intempestifs

Public gSessionInitialisee As Boolean

Private Const NOM_FICHIER_TRACE As String = "trace_session.txt"
Private Const NOM_FICHIER_LOG As String = "log_session.txt"

Public Sub CreerTraceOuverture()

    On Error Resume Next
        Dim chemin As String: chemin = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
        Dim f As Integer: f = FreeFile
        Open chemin & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & "" & NOM_FICHIER_TRACE For Output As #f
        Print #f, "Ouverture normale à " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        Close #f
    On Error GoTo 0

    'Cellule témoin
    On Error Resume Next
        wsdADMIN.Range("Z1").Value = "Ouvert à " & Format(Now, "hh:nn:ss")
    On Error GoTo 0
    
End Sub

Public Sub SupprimerTraceOuverture()

    On Error Resume Next
        Dim chemin As String: chemin = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
        Kill chemin & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & "" & NOM_FICHIER_TRACE
    On Error GoTo 0
    
End Sub

Public Sub VerifierOuvertureSilencieuse()

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modTraceSession:VerifierOuvertureSilencieuse", vbNullString, 0)
    
    On Error Resume Next
        Dim chemin As String: chemin = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
        If Dir(chemin & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & "" & NOM_FICHIER_TRACE) <> vbNullString Then
            If Timer - FileDateTime(chemin & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & "" & NOM_FICHIER_TRACE) > 5 Then
                'Le fichier existe mais Workbook_Open n’a pas été exécuté
                Call AppelLogSessionAnormale("Ouverture silencieuse détectée — Workbook_Open non exécuté")
            End If
        End If
    On Error GoTo 0
    
    Call modDev_Utils.EnregistrerLogApplication("modTraceSession:VerifierOuvertureSilencieuse", vbNullString, startTime)

End Sub

Private Sub AppelLogSessionAnormale(msg As String)

    On Error Resume Next
        Dim chemin As String: chemin = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
        Dim f As Integer: f = FreeFile
        Open chemin & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & "" & NOM_FICHIER_LOG For Append As #f
        Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & Fn_UtilisateurWindows & " | " & msg
        Close #f
    On Error GoTo 0
    
End Sub

