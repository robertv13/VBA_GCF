Attribute VB_Name = "modLog_Analysis"
Option Explicit

Sub Lire_Fichier_LogMainApp() '2025-01-10 @ 17:11

    'Initialiser la feuille où les données seront insérées
    
    Dim ws As Worksheet: Set ws = wshzDocLogMainAppAnalysis
    
    'Utiliser une boîte de dialogue Fichier
    Dim FileDialogBox As FileDialog
    Set FileDialogBox = Application.FileDialog(msoFileDialogFilePicker)

    'Configurer les filtres
    Dim FilePath As String
    With FileDialogBox
        .Title = "Sélectionnez le fichier LogMainApp à analyser"
        .Filters.Clear 'Supprimer les filtres existants
        .Filters.Add "Fichiers log", "*.log"
        .Filters.Add "Tous les fichiers", "*.*"

        'Afficher la boîte de dialogue
        If .show = -1 Then
            FilePath = .selectedItems(1) 'Récupérer le chemin du fichier sélectionné
        Else
            MsgBox "Aucun fichier sélectionné.", vbExclamation
            Exit Sub
        End If
    End With

    'Détermine l'environnement (DEV/PROD) ?
    Dim env As String
    If Not InStr(FilePath, "C:\VBA\GC_FISCALITÉ\DataFiles\") = 1 Then
        env = "PROD"
    Else
        env = "DEV"
    End If
    
    'Ouvrir le fichier sélectionné
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open FilePath For Input As #FileNumber

    'Utilisation d'un tableau pour préparer les données
    Dim output() As Variant
    ReDim output(1 To 50000, 1 To 8)
    
    'Lire le fichier ligne par ligne et traiter les données
    Dim ligne As Long: ligne = 1
    Dim LineContent As String
    Dim lineNumber As Long
    Dim Fields() As String
    Dim duree As String
    Dim i As Long

    Do While Not EOF(FileNumber)
        Line Input #FileNumber, LineContent
        lineNumber = lineNumber + 1
        If Trim(LineContent) <> "" Then
            If InStr(LineContent, " | ") <> 0 Then
                Fields = Split(LineContent, " | ") 'Diviser la ligne en champs avec le délimiteur "|"
                'Insérer les données dans la feuille
                output(ligne, 1) = env
                output(ligne, 2) = CStr(Left(Fields(0), 10))
                output(ligne, 3) = CStr(Right(Fields(0), 11))
                output(ligne, 4) = Trim(Fields(1))
                output(ligne, 5) = Trim(Fields(2))
                output(ligne, 6) = Trim(Fields(3))
                If InStr(Fields(3), " secondes'") <> 0 Then
                    duree = ExtraireSecondes(Fields(3))
                    duree = Replace(duree, ".", ",")
'                    duree = Mid(Fields(3), InStr(Fields(3), " *** = '") + 8)
'                    duree = Left(duree, InStr(duree, " ") - 1)
                    output(ligne, 7) = CDbl(duree)
                    output(ligne, 6) = Trim(Left(Fields(3), InStr(Fields(3), " = ") - 1)) & " (S)"
                End If
                output(ligne, 8) = lineNumber
                ligne = ligne + 1
            End If
        End If
    Loop

    ' Fermer le fichier
    Close #FileNumber
    
    Call Array_2D_Resizer(output, ligne, UBound(output, 2))
    
    'Ajout du tableau (output) en une seule opération après ce qui existe déjà
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ws.Cells(lastUsedRow + 1, 1).Resize(UBound(output, 1), UBound(output, 2)).Value = output
    
'    Dim rng As Range
'    Set rng = ws.Range("A" & lastUsedRow + 1).Resize(UBound(output, 1), UBound(output, 2))
'    rng.Value = output
    
    'Appliquer le format de date à la première colonne
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim rng As Range
    Set rng = ws.Range("A1:H" & lastUsedRow)
    rng.Columns(2).NumberFormat = "yyyy-mm-dd"
    rng.Columns(3).NumberFormat = "hh:mm:ss.00"
    
    ws.Range("G:G").NumberFormat = "##0.0000"
    ws.Range("G:G").HorizontalAlignment = xlCenter

    'Libérer la mémoire
    Set FileDialogBox = Nothing
    Set rng = Nothing
    Set ws = Nothing

    MsgBox "Lecture du fichier terminée et données insérées dans la feuille.", vbInformation

End Sub

