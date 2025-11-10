Attribute VB_Name = "modSessionVerrou"
Option Explicit

Public Sub VerrouillerSiSessionInvalide(Optional contexte As String = "Interaction") '2025-11-10 @ 10:50

    If Not SessionEstValide() Then
        MsgBox "Session invalide détectée : " & contexte & vbNewLine & vbNewLine & _
               "Veuillez relancer l'application via le raccourci prévu.", _
               vbCritical
        ThisWorkbook.Close SaveChanges:=False
    End If
    
    
End Sub

Public Function SessionEstValide() As Boolean '2025-11-10 @ 10:50

    Dim chemin As String: chemin = Fn_RepertoireBaseApplication(Fn_UtilisateurWindows)
    SessionEstValide = (Dir(Fn_RepertoireBaseApplication(Fn_UtilisateurWindows) & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & "trace_session.txt") <> vbNullString)
    
End Function

