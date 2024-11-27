Attribute VB_Name = "modLog_Analysis"
Option Explicit

Sub Lire_Fichier_LogMainApp()

    ' Initialiser la feuille o� les donn�es seront ins�r�es
    Dim ws As Worksheet: Set ws = wshzDocLogMainAppAnalysis
    Dim ligne As Long
    ligne = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    ws.Range("A1:D" & ligne).Clear
    ws.Cells(1, 1).value = "Date"
    ws.Cells(1, 2).value = "Heure"
    ws.Cells(1, 3).value = "Utilisateur"
    ws.Cells(1, 4).value = "VersionApp"
    ws.Cells(1, 5).value = "Commentaires"
    ws.Cells(1, 6).value = "Secondes"
    ws.Range("A:E").NumberFormat = "@"
    ws.Range("F:F").NumberFormat = "##0.0000"
    ligne = 2

    'Cr�er une bo�te de dialogue Fichier
    Dim FileDialogBox As FileDialog
    Set FileDialogBox = Application.FileDialog(msoFileDialogFilePicker)

    'Configurer les filtres
    Dim FilePath As String
    With FileDialogBox
        .Title = "S�lectionnez le fichier LogMainApp"
        .Filters.Clear 'Supprimer les filtres existants
'        .Filters.Add "Fichiers texte", "*.txt"
        .Filters.add "Fichiers log", "*.log"
        .Filters.add "Tous les fichiers", "*.*"

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
    Dim Fields() As String
    Dim duree As String
    Dim i As Long

    Do While Not EOF(FileNumber)
        Line Input #FileNumber, LineContent
        If InStr(LineContent, "|") <> 0 Then
            Fields = Split(LineContent, "|") ' Diviser la ligne en champs avec le d�limiteur "|"
            'Ins�rer les donn�es dans la feuille
            ws.Cells(ligne, 1) = Left(Fields(0), 10)
            ws.Cells(ligne, 2) = Right(Fields(0), 12)
            ws.Cells(ligne, 3) = Trim(Fields(1))
            ws.Cells(ligne, 4) = Trim(Fields(2))
            ws.Cells(ligne, 5) = Trim(Fields(3))
            If InStr(Fields(3), " = ") <> 0 Then
                duree = Mid(Fields(3), InStr(Fields(3), " = ") + 3, Len(Fields(3)) - InStr(Fields(3), " = "))
                ws.Cells(ligne, 6).value = CDbl(Left(duree, InStr(duree, " ") - 1))
                ws.Cells(ligne, 5).value = Trim(Left(Fields(3), InStr(Fields(3), " = ") - 1))
            End If
            ligne = ligne + 1
        End If
    Loop

    ' Fermer le fichier
    Close #FileNumber

    'Lib�rer la m�moire
    Set FileDialogBox = Nothing
    Set ws = Nothing
    
    MsgBox "Lecture du fichier termin�e et donn�es ins�r�es dans la feuille.", vbInformation
    
End Sub

'Sub Lire_Fichier_LogMainApp()
'
'    Dim cheminFichier As String
'    Dim fichierLOG As Integer
'    Dim ligneTexte As String
'    Dim champs() As String
'    Dim data() As Variant
'    Dim ligneNum As Long
'    Dim maxColonnes As Long
'    Dim dlg As FileDialog
'
'    ' S�lectionner le fichier avec FileDialog
'    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
'    With dlg
'        .Title = "S�lectionnez l'emplacement du fichier LogMainApp.log"
'        .AllowMultiSelect = False
'        .Filters.Clear
'        .Filters.Add "Fichiers TXT", "*.txt"
'        .Filters.Add "Fichiers LOG", "*.log"
'        .Filters.Add "Tous les fichiers", "*.*"
'
'        If .show = -1 Then
'            cheminFichier = .selectedItems(1)
'        Else
'            MsgBox "Aucun fichier LogMainApp de s�lectionn�."
'            Exit Sub
'        End If
'    End With
'
'    'Initialisation de la lecture du fichier
'    fichierLOG = FreeFile
'    On Error GoTo ErreurOuverture
'    Open cheminFichier For Input As #fichierLOG
'
'    'Premi�re passe : d�terminer le nombre de lignes et de colonnes
'    ligneNum = 0
'    Do While Not EOF(fichierLOG)
'        Line Input #fichierLOG, ligneTexte
'        champs = Split(ligneTexte, " | ")
'        maxColonnes = Application.Max(maxColonnes, UBound(champs) + 1)
'        ligneNum = ligneNum + 1
'    Loop
'
'    'Dimensionner le tableau 2D pour toutes les donn�es
'    ReDim data(1 To ligneNum, 1 To maxColonnes)
'
'    'Revenir au d�but du fichier pour la deuxi�me passe
'    Close #fichierLOG
'    Open cheminFichier For Input As #fichierLOG
'
'    'Deuxi�me passe : remplir le tableau avec les donn�es
'    Dim i As Long, j As Long
'    i = 1
'    Do While Not EOF(fichierLOG)
'        Line Input #fichierLOG, ligneTexte
'        champs = Split(ligneTexte, " | ")
'        For j = LBound(champs) To UBound(champs)
'            data(i, j + 1) = champs(j)
'        Next j
'        i = i + 1
'    Loop
'
'    Close #fichierLOG
'    MsgBox "Donn�es charg�es dans le tableau."
'
'    'Exemple de traitement : afficher le contenu du tableau dans la fen�tre d'ex�cution
'    Dim moment As String
'    Dim user As String
'    Dim version As String
'    Dim procedure As String
'
'    Dim nbEntree As Long
'
'    Dim dicMoment As Dictionary: Set dicMoment = New Dictionary
'    Dim dicUser As Dictionary: Set dicUser = New Dictionary
'    For i = 1 To UBound(data, 1)
'        If Trim(data(i, 1)) <> "" Then
'            nbEntree = nbEntree + 1
'
'            moment = Left(data(i, 1), 13)
'            If Not dicMoment.Exists(moment) Then
'                dicMoment.Add moment, 0
'            End If
'            dicMoment.item(moment) = dicMoment.item(moment) + 1
'
'            user = Trim(data(i, 2))
'            If Not dicUser.Exists(user) Then
'                dicUser.Add user, 0
'            End If
'            dicUser.item(user) = dicUser.item(user) + 1
'
'            version = Trim(data(i, 3))
'            procedure = Trim(data(i, 4))
'        End If
'    Next i
'
'    'Quand l'application est-elle utilis�e ?
'    Dim cles() As Variant
'    cles = dicMoment.keys  'R�cup�rer toutes les cl�s du dictionnaire dans un tableau
'
'    'Trier le tableau de cl�s
'    Dim temp As Variant
'    For i = LBound(cles) To UBound(cles) - 1
'        For j = i + 1 To UBound(cles)
'            If cles(i) > cles(j) Then
'                '�changer les �l�ments pour trier
'                temp = cles(i)
'                cles(i) = cles(j)
'                cles(j) = temp
'            End If
'        Next j
'    Next i
'
'    'Afficher chaque combinaison cl�/valeur dans la fen�tre d'ex�cution
'    Debug.Print "#070 - " & vbNewLine & "Quand l'application est utilis�e ?"
'    For i = LBound(cles) To UBound(cles)
'        Debug.Print "#071 - " & Space(5); cles(i); Tab(22); " soit " & dicMoment(cles(i)) & " entr�es "; Tab(42); "ou " & Format$(dicMoment(cles(i)) / nbEntree, "##0.00 %")
'    Next i
'
'    'Qui utilise l'application ?
'    cles = dicUser.keys  'R�cup�rer toutes les cl�s du dictionnaire dans un tableau
'
'    'Trier le tableau de cl�s
'    For i = LBound(cles) To UBound(cles) - 1
'        For j = i + 1 To UBound(cles)
'            If cles(i) > cles(j) Then
'                '�changer les �l�ments pour trier
'                temp = cles(i)
'                cles(i) = cles(j)
'                cles(j) = temp
'            End If
'        Next j
'    Next i
'
'    'Afficher chaque combinaison cl�/valeur dans la fen�tre d'ex�cution
'    Debug.Print "#072 - " & vbNewLine & "Qui utilise l'application ?"
'    For i = LBound(cles) To UBound(cles)
'        Debug.Print "#073 - " & Space(5); cles(i); Tab(22); " soit " & dicUser(cles(i)) & " entr�es "; Tab(42); "ou " & Format$(dicUser(cles(i)) / nbEntree, "##0.00 %")
'    Next i
'
'    MsgBox "Le traitement est termin�"
'
'    Exit Sub
'
'ErreurOuverture:
'    MsgBox "Erreur lors de l'ouverture du fichier : " & Err.description
'End Sub

