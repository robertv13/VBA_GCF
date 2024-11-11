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

    ' S�lectionner le fichier avec FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "S�lectionnez l'emplacement du fichier LogMainApp.txt"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Fichiers LOG", "*.txt", 1
        .Filters.Add "Tous les fichiers", "*.*"
        
        If .show = -1 Then
            cheminFichier = .selectedItems(1)
        Else
            MsgBox "Aucun fichier LogMainApp de s�lectionn�."
            Exit Sub
        End If
    End With
    
    'Initialisation de la lecture du fichier
    fichierLOG = FreeFile
    On Error GoTo ErreurOuverture
    Open cheminFichier For Input As #fichierLOG

    'Premi�re passe : d�terminer le nombre de lignes et de colonnes
    ligneNum = 0
    Do While Not EOF(fichierLOG)
        Line Input #fichierLOG, ligneTexte
        champs = Split(ligneTexte, " | ")
        maxColonnes = Application.Max(maxColonnes, UBound(champs) + 1)
        ligneNum = ligneNum + 1
    Loop

    'Dimensionner le tableau 2D pour toutes les donn�es
    ReDim data(1 To ligneNum, 1 To maxColonnes)
    
    'Revenir au d�but du fichier pour la deuxi�me passe
    Close #fichierLOG
    Open cheminFichier For Input As #fichierLOG
    
    'Deuxi�me passe : remplir le tableau avec les donn�es
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
    MsgBox "Donn�es charg�es dans le tableau."

    'Exemple de traitement : afficher le contenu du tableau dans la fen�tre d'ex�cution
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

    'Quand l'application est-elle utilis�e ?
    Dim cles() As Variant
    cles = dicMoment.keys  'R�cup�rer toutes les cl�s du dictionnaire dans un tableau
    
    'Trier le tableau de cl�s
    Dim temp As Variant
    For i = LBound(cles) To UBound(cles) - 1
        For j = i + 1 To UBound(cles)
            If cles(i) > cles(j) Then
                '�changer les �l�ments pour trier
                temp = cles(i)
                cles(i) = cles(j)
                cles(j) = temp
            End If
        Next j
    Next i
    
    'Afficher chaque combinaison cl�/valeur dans la fen�tre d'ex�cution
    Debug.Print vbNewLine & "Quand l'application est utilis�e ?"
    For i = LBound(cles) To UBound(cles)
        Debug.Print Space(5); cles(i); Tab(22); " soit " & dicMoment(cles(i)) & " entr�es "; Tab(42); "ou " & Format$(dicMoment(cles(i)) / nbEntree, "##0.00 %")
    Next i
    
    'Qui utilise l'application ?
    cles = dicUser.keys  'R�cup�rer toutes les cl�s du dictionnaire dans un tableau
    
    'Trier le tableau de cl�s
    For i = LBound(cles) To UBound(cles) - 1
        For j = i + 1 To UBound(cles)
            If cles(i) > cles(j) Then
                '�changer les �l�ments pour trier
                temp = cles(i)
                cles(i) = cles(j)
                cles(j) = temp
            End If
        Next j
    Next i
    
    'Afficher chaque combinaison cl�/valeur dans la fen�tre d'ex�cution
    Debug.Print vbNewLine & "Qui utilise l'application ?"
    For i = LBound(cles) To UBound(cles)
        Debug.Print Space(5); cles(i); Tab(22); " soit " & dicUser(cles(i)) & " entr�es "; Tab(42); "ou " & Format$(dicUser(cles(i)) / nbEntree, "##0.00 %")
    Next i
    
    MsgBox "Le traitement est termin�"
    
    Exit Sub

ErreurOuverture:
    MsgBox "Erreur lors de l'ouverture du fichier : " & Err.description
End Sub

