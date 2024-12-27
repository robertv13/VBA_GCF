Attribute VB_Name = "modLog_Analysis"
Option Explicit

Sub Lire_Fichier_LogMainApp()

    'Initialiser la feuille o� les donn�es seront ins�r�es
    
    Dim ws As Worksheet: Set ws = wshzDocLogMainAppAnalysis
    ws.Range("A1").CurrentRegion.Clear
    
    Dim output() As Variant
    ReDim output(1 To 5000, 1 To 7)
    
    output(1, 1) = "Date"
    output(1, 2) = "Heure"
    output(1, 3) = "Utilisateur"
    output(1, 4) = "VersionApp"
    output(1, 5) = "Commentaires"
    output(1, 6) = "Secondes"
    
    Dim ligne As Long
    ligne = 2

    'Cr�er une bo�te de dialogue Fichier
    Dim FileDialogBox As FileDialog
    Set FileDialogBox = Application.FileDialog(msoFileDialogFilePicker)

    'Configurer les filtres
    Dim FilePath As String
    With FileDialogBox
        .Title = "S�lectionnez le fichier LogMainApp"
        .Filters.Clear 'Supprimer les filtres existants
'        .Filters.add "Fichiers texte", "*.txt"
        .Filters.Add "Fichiers log", "*.log"
        .Filters.Add "Tous les fichiers", "*.*"

        'Afficher la bo�te de dialogue
        If .show = -1 Then
            FilePath = .selectedItems(1) ' R�cup�rer le chemin du fichier s�lectionn�
        Else
            MsgBox "Aucun fichier s�lectionn�.", vbExclamation
            Exit Sub
        End If
    End With

    ' Ouvrir le fichier s�lectionn�
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open FilePath For Input As #FileNumber

    'Lire le fichier ligne par ligne et traiter les donn�es
    Dim LineContent As String
    Dim lineNumber As Long
    Dim Fields() As String
    Dim duree As String
    Dim i As Long

    Do While Not EOF(FileNumber)
        Line Input #FileNumber, LineContent
        lineNumber = lineNumber + 1
        If Trim(LineContent) <> "" Then
    '        If InStr(LineContent, "modENC_Saisie:MAJ_Encaissement") <> 0 Then Stop
            If InStr(LineContent, "|") <> 0 Then
                Fields = Split(LineContent, "|") ' Diviser la ligne en champs avec le d�limiteur "|"
                'Ins�rer les donn�es dans la feuille
                output(ligne, 1) = CStr(Left(Fields(0), 10))
                output(ligne, 2) = CStr(Right(Fields(0), 12))
                output(ligne, 3) = Trim(Fields(1))
                output(ligne, 4) = Trim(Fields(2))
                output(ligne, 5) = Trim(Fields(3))
                If InStr(Fields(3), " = ") <> 0 Then
                    duree = Mid(Fields(3), InStr(Fields(3), " = ") + 4, Len(Fields(3)) - InStr(Fields(3), " = "))
                    output(ligne, 6) = CDbl(Left(duree, InStr(duree, " ") - 1))
                    output(ligne, 5) = Trim(Left(Fields(3), InStr(Fields(3), " = ") - 1)) & " (S)"
                    If output(ligne, 6) > 25 Then Stop
                End If
                output(ligne, 7) = lineNumber
                ligne = ligne + 1
            End If
        End If
    Loop

    ' Fermer le fichier
    Close #FileNumber
    
    Call Array_2D_Resizer(output, ligne, UBound(output, 2))
    
    Dim rng As Range
    Set rng = ws.Range("A1").Resize(UBound(output, 1), UBound(output, 2))
    rng.Value = output
    
    'Appliquer le format de date � la premi�re colonne
    rng.Columns(1).NumberFormat = "yyyy-mm-dd"
    rng.Columns(2).NumberFormat = "hh:mm:ss.00"
    
    ws.Range("F:F").NumberFormat = "##0.0000"
    ws.Range("F:F").HorizontalAlignment = xlRight

    'Lib�rer la m�moire
    Set FileDialogBox = Nothing
    Set rng = Nothing
    Set ws = Nothing

    MsgBox "Lecture du fichier termin�e et donn�es ins�r�es dans la feuille.", vbInformation

End Sub

