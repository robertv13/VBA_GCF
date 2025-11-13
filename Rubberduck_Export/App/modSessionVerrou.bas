Attribute VB_Name = "modSessionVerrou"
Option Explicit

Public Sub VerrouillerSiSessionInvalide(Optional contexte As String = "Interaction") '2025-11-10 @ 10:50

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modSessionVerrou:VerrouillerSiSessionInvalide", vbNullString, 0)

    If Not SessionEstValideComplet() Then
        Call EnregistrerErreurs("modSessionVerrou", "VerrouillerSiSessionInvalide", _
                        "trace_session.txt manquant -ET/OU- Actif_.txt manquant", Err.Number, "ERREUR")
        MsgBox "La présente session est INVALIDE" & vbNewLine & vbNewLine & _
               contexte & vbNewLine & vbNewLine & _
               "Veuillez relancer l'application via le raccourci prévu.", _
               vbCritical, _
               "VerrouillerSiSessionInvalide"
        Call FermerApplication("Session invalide détectée", True)
    End If
    
    Call modDev_Utils.EnregistrerLogApplication("modSessionVerrou:VerrouillerSiSessionInvalide", _
                                                            vbNullString, startTime)

End Sub

Public Function SessionEstValideComplet() As Boolean '2025-11-11 @ 07:35

    Dim basePath As String: basePath = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
    
    Dim traceOK As Boolean
    traceOK = (Dir(basePath & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & _
              "trace_session.txt") <> vbNullString)
    
    Dim actifOK As Boolean
    actifOK = (Dir(basePath & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & _
              "Actif_" & Fn_UtilisateurWindows & ".txt") <> vbNullString)
    
    SessionEstValideComplet = gSessionInitialisee And traceOK And actifOK
    
    If Fn_UtilisateurWindows = "RobertMV" And gSessionInitialisee = False Then gSessionInitialisee = True '2025-11-12 @ 18:54
    
    If SessionEstValideComplet = False Then
        Debug.Print "gSessionInitialisee = " & gSessionInitialisee & "    traceOK = " & traceOK & _
                                                    "     actifOK = " & actifOK & "     ALORS = " & _
                                                    (gSessionInitialisee And traceOK And actifOK)
    End If
    
End Function

