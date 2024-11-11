Attribute VB_Name = "modLog_Analysis"
Option Explicit

Sub Lire_Fichier_LogMainApp()
    
    Dim cheminFichier As String
    Dim fichierLOG As Integer
    Dim ligneTexte As String
    Dim champs() As String
    Dim data() As Variant
    Dim ligneNum As Long
    Dim maxColonnes As Long
    Dim dlg As FileDialog

    ' Sélectionner le fichier avec FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "Sélectionnez l'emplacement du fichier LogMainApp.txt"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Fichiers LOG", "*.txt", 1
        .Filters.Add "Tous les fichiers", "*.*"
        
        If .show = -1 Then
            cheminFichier = .selectedItems(1)
        Else
            MsgBox "Aucun fichier LogMainApp de sélectionné."
            Exit Sub
        End If
    End With
    
    'Initialisation de la lecture du fichier
    fichierLOG = FreeFile
    On Error GoTo ErreurOuverture
    Open cheminFichier For Input As #fichierLOG

    'Première passe : déterminer le nombre de lignes et de colonnes
    ligneNum = 0
    Do While Not EOF(fichierLOG)
        Line Input #fichierLOG, ligneTexte
        champs = Split(ligneTexte, " | ")
        maxColonnes = Application.Max(maxColonnes, UBound(champs) + 1)
        ligneNum = ligneNum + 1
    Loop

    'Dimensionner le tableau 2D pour toutes les données
    ReDim data(1 To ligneNum, 1 To maxColonnes)
    
    'Revenir au début du fichier pour la deuxième passe
    Close #fichierLOG
    Open cheminFichier For Input As #fichierLOG
    
    'Deuxième passe : remplir le tableau avec les données
    Dim i As Long, j As Long
    i = 1
    Do While Not EOF(fichierLOG)
        Line Input #fichierLOG, ligneTexte
        champs = Split(ligneTexte, " | ")
        For j = LBound(champs) To UBound(champs)
            data(i, j + 1) = champs(j)
        Next j
        i = i + 1
    Loop

    Close #fichierLOG
    MsgBox "Données chargées dans le tableau."

    'Exemple de traitement : afficher le contenu du tableau dans la fenêtre d'exécution
    Dim moment As String
    Dim user As String
    Dim version As String
    Dim procedure As String
    
    Dim nbEntree As Long
    
    Dim dicMoment As Dictionary: Set dicMoment = New Dictionary
    Dim dicUser As Dictionary: Set dicUser = New Dictionary
    For i = 1 To UBound(data, 1)
        If Trim(data(i, 1)) <> "" Then
            nbEntree = nbEntree + 1
            
            moment = Left(data(i, 1), 13)
            If Not dicMoment.Exists(moment) Then
                dicMoment.Add moment, 0
            End If
            dicMoment.item(moment) = dicMoment.item(moment) + 1
            
            user = Trim(data(i, 2))
            If Not dicUser.Exists(user) Then
                dicUser.Add user, 0
            End If
            dicUser.item(user) = dicUser.item(user) + 1
            
            version = Trim(data(i, 3))
            procedure = Trim(data(i, 4))
        End If
    Next i

    'Quand l'application est-elle utilisée ?
    Dim cles() As Variant
    cles = dicMoment.keys  'Récupérer toutes les clés du dictionnaire dans un tableau
    
    'Trier le tableau de clés
    Dim temp As Variant
    For i = LBound(cles) To UBound(cles) - 1
        For j = i + 1 To UBound(cles)
            If cles(i) > cles(j) Then
                'Échanger les éléments pour trier
                temp = cles(i)
                cles(i) = cles(j)
                cles(j) = temp
            End If
        Next j
    Next i
    
    'Afficher chaque combinaison clé/valeur dans la fenêtre d'exécution
    Debug.Print vbNewLine & "Quand l'application est utilisée ?"
    For i = LBound(cles) To UBound(cles)
        Debug.Print Space(5); cles(i); Tab(22); " soit " & dicMoment(cles(i)) & " entrées "; Tab(42); "ou " & Format$(dicMoment(cles(i)) / nbEntree, "##0.00 %")
    Next i
    
    'Qui utilise l'application ?
    cles = dicUser.keys  'Récupérer toutes les clés du dictionnaire dans un tableau
    
    'Trier le tableau de clés
    For i = LBound(cles) To UBound(cles) - 1
        For j = i + 1 To UBound(cles)
            If cles(i) > cles(j) Then
                'Échanger les éléments pour trier
                temp = cles(i)
                cles(i) = cles(j)
                cles(j) = temp
            End If
        Next j
    Next i
    
    'Afficher chaque combinaison clé/valeur dans la fenêtre d'exécution
    Debug.Print vbNewLine & "Qui utilise l'application ?"
    For i = LBound(cles) To UBound(cles)
        Debug.Print Space(5); cles(i); Tab(22); " soit " & dicUser(cles(i)) & " entrées "; Tab(42); "ou " & Format$(dicUser(cles(i)) / nbEntree, "##0.00 %")
    Next i
    
    MsgBox "Le traitement est terminé"
    
    Exit Sub

ErreurOuverture:
    MsgBox "Erreur lors de l'ouverture du fichier : " & Err.description
End Sub

