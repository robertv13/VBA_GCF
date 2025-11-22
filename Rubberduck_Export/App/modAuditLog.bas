Attribute VB_Name = "modAuditLog"
Option Explicit

Public Sub VerifierValiditeSessions(wsOutput As Worksheet, _
                                                                    ByRef r As Long, logPath As String)

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.fileExists(logPath) Then
        Call ReporterAnomalie(wsOutput, r, "Je n'ai trouvé AUCUN fichier .log dans " & logPath, _
                                                                            vbNullString, vbNullString)
        r = r + 1
        Exit Sub
    End If

    Dim ts As Object: Set ts = fso.OpenTextFile(logPath, 1)
    Dim lignes() As String: lignes = Split(ts.ReadAll, vbNewLine)
    ts.Close

    '--- Règle 2 : première entrée du fichier ---
    Dim firstLine As String
    firstLine = Trim(lignes(0))
    If firstLine <> "" Then
        If InStr(firstLine, "DÉBUT D'UNE NOUVELLE SESSION") = 0 Then
            Call ReporterAnomalie(wsOutput, r, Space(5) & "Première entrée est INVALIDE", _
                                               "L=" & CStr(1) & " " & firstLine, _
                                               vbNullString)
        End If
    End If

    Dim fermetureEnAttente As Object: Set fermetureEnAttente = CreateObject("Scripting.Dictionary")
    Dim fermetureInfos As Object: Set fermetureInfos = CreateObject("Scripting.Dictionary")
    Dim utilisateursVus As Object: Set utilisateursVus = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 0 To UBound(lignes)
        Dim ligne As String: ligne = Trim(lignes(i))
        If ligne = "" Then GoTo Suivant

        Dim champs() As String: champs = Split(ligne, " | ")
        If UBound(champs) < 3 Then GoTo Suivant

        Dim horodatage As String: horodatage = champs(0)
        Dim utilisateur As String: utilisateur = champs(1)
        Dim description As String: description = champs(3)

        '--- Règle 3 : première apparition utilisateur ---
        If Not utilisateursVus.Exists(utilisateur) Then
            If InStr(description, "DÉBUT D'UNE NOUVELLE SESSION") = 0 Then
                Call ReporterAnomalie(wsOutput, r, Space(5) & "Première entrée est non conforme", _
                                                   "L=" & CStr(i) & " " & ligne, _
                                                   vbNullString)
            End If
            utilisateursVus(utilisateur) = True
        End If

        '--- Fermeture ---
        If InStr(UCase(description), "SESSION TERMINÉE") > 0 Then
            fermetureEnAttente(utilisateur) = True
            fermetureInfos(utilisateur & "_ligneTexte") = ligne
            GoTo Suivant
        End If

        '--- Démarrage explicite ---
        If InStr(description, "DÉBUT D'UNE NOUVELLE SESSION") > 0 Then
            If fermetureEnAttente.Exists(utilisateur) Then
                fermetureEnAttente.Remove utilisateur
                fermetureInfos.Remove utilisateur & "_ligneTexte"
            End If
            GoTo Suivant
        End If

        ' --- Règle 1 : réapparition sans redémarrage ---
        If fermetureEnAttente.Exists(utilisateur) Then
            Call ReporterAnomalie(wsOutput, r, Space(5) & "Réapparition sans Workbook_Open", _
                                               "L=" & CStr(i) & " " & ligne, _
                                               "Après fermeture : " & fermetureInfos(utilisateur & "_ligneTexte"))
            fermetureEnAttente.Remove utilisateur
            fermetureInfos.Remove utilisateur & "_ligneTexte"
        End If

Suivant:
    Next i
    
    r = r + 1

End Sub

