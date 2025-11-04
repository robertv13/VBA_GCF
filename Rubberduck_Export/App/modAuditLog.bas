Attribute VB_Name = "modAuditLog"
Option Explicit

Sub MAIN() '2025-11-04 @ 09:33

    Dim logPath As String
    ChDrive "C:\VBA\GC_FISCALITÉ"
    ChDir "C:\VBA\GC_FISCALITÉ"
    logPath = Application.GetOpenFilename("Fichiers texte (*.txt; *.log), *.txt; *.log", , "Fichier log à analyser")

    Call AuditerSessionsAvecContexteEtReapparition(logPath)
    
End Sub

Sub AuditerSessionsAvecContexteEtReapparition(logPath As String)

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.fileExists(logPath) Then
        MsgBox "Fichier log introuvable : " & logPath, vbCritical
        Exit Sub
    End If

    Dim ts As Object: Set ts = fso.OpenTextFile(logPath, 1)
    Dim lignes() As String: lignes = Split(ts.ReadAll, vbNewLine)
    ts.Close

    Debug.Print vbCrLf
    Dim parts As Variant
    parts = Split(logPath, "\")
    Debug.Print parts(4) & vbCrLf & "---------------"
    Dim i As Long
    Dim fermetureEnAttente As Object: Set fermetureEnAttente = CreateObject("Scripting.Dictionary")
    Dim fermetureInfos As Object: Set fermetureInfos = CreateObject("Scripting.Dictionary")

    For i = 0 To UBound(lignes)
        Dim ligne As String: ligne = Trim(lignes(i))
        If ligne = "" Then GoTo Suivant

        Dim champs() As String: champs = Split(ligne, " | ")
        If UBound(champs) < 3 Then GoTo Suivant

        Dim horodatage As String: horodatage = champs(0)
        Dim utilisateur As String: utilisateur = champs(1)
        Dim description As String: description = champs(3)

        ' --- Fin de session normale ---
        If InStr(description, "Session terminée NORMALEMENT") > 0 Then
            fermetureEnAttente(utilisateur) = True
            fermetureInfos(utilisateur & "_ligne") = i
            fermetureInfos(utilisateur & "_horodatage") = horodatage
            fermetureInfos(utilisateur & "_ligneTexte") = ligne
            GoTo Suivant
        End If

        ' --- Démarrage explicite ---
        If InStr(description, "DÉBUT D'UNE NOUVELLE SESSION") > 0 Then
            If fermetureEnAttente.Exists(utilisateur) Then
                fermetureEnAttente.Remove utilisateur
                fermetureInfos.Remove utilisateur & "_ligne"
                fermetureInfos.Remove utilisateur & "_horodatage"
                fermetureInfos.Remove utilisateur & "_ligneTexte"
            End If
            GoTo Suivant
        End If

        ' --- Réapparition sans redémarrage ---
        If fermetureEnAttente.Exists(utilisateur) Then
            Dim ligneFermeture As Long: ligneFermeture = fermetureInfos(utilisateur & "_ligne")
            Dim horoFermeture As Date
            Dim horoTexte As String: horoTexte = fermetureInfos(utilisateur & "_horodatage")
            If InStr(horoTexte, ".") > 0 Then
                horoTexte = Left(horoTexte, InStr(horoTexte, ".") - 1) ' supprimer les centièmes de seconde
            End If
            
            If IsDate(horoTexte) Then
                horoFermeture = CDate(horoTexte)
            Else
                Debug.Print "?? Format de date non reconnu : " & horoTexte
                GoTo Suivant
            End If
            horodatage = Left(horodatage, 19)
            Dim horoActuel As Date: horoActuel = CDate(horodatage)
            Dim delta As Double: delta = (horoActuel - horoFermeture) * 86400 ' en secondes

            Debug.Print "Réapparition sans redémarrage pour '" & utilisateur & "'"
            Debug.Print "Fermeture (ligne " & ligneFermeture & ") : " & fermetureInfos(utilisateur & "_ligneTexte")
            Debug.Print "Réapparition (ligne " & i & ") : " & ligne
            Debug.Print "Temps écoulé : " & Fn_FormatDuree(delta)
            Debug.Print String(60, "-")

            fermetureEnAttente.Remove utilisateur
            fermetureInfos.Remove utilisateur & "_ligne"
            fermetureInfos.Remove utilisateur & "_horodatage"
            fermetureInfos.Remove utilisateur & "_ligneTexte"
        End If

Suivant:
    Next i

    MsgBox "Audit terminé : voir la fenêtre Exécution (Ctrl+G)", vbInformation

End Sub

Function Fn_FormatDuree(deltaSeconds As Double) As String

    Dim heures As Long, minutes As Long, secondes As Long

    heures = Int(deltaSeconds \ 3600)
    minutes = Int((deltaSeconds Mod 3600) \ 60)
    secondes = Int(deltaSeconds Mod 60)

    Fn_FormatDuree = heures & "h " & minutes & "m " & secondes & "s"
    
End Function

