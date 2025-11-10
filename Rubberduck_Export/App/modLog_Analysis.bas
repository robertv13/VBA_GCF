Attribute VB_Name = "modLog_Analysis"
Option Explicit

Sub shpTraiterLogFiles_Click()

    Call OuvrirRepertoireLogEtTraiterFichiers

End Sub

Sub OuvrirRepertoireLogEtTraiterFichiers()

    'Initialisation du FileDialog pour sélectionner un répertoire
    Dim fileDialog As fileDialog
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    With fileDialog
        .Title = "Sélectionnez un répertoire à traiter"
        .InitialFileName = "C:\VBA\GC_FISCALITÉ\GCF_DataFiles\"
    End With
    
    'Un répertoire a-t-il été sélectionné ?
    Dim folderPath As String
    If fileDialog.show = -1 Then
        folderPath = fileDialog.SelectedItems(1)
    Else
        MsgBox "Aucun répertoire sélectionné.", vbExclamation
        Exit Sub
    End If
    
    'Vérification de l'existence du répertoire
    If Dir(folderPath, vbDirectory) = vbNullString Then
        MsgBox "Répertoire invalide.", vbCritical
        Exit Sub
    End If
    
    'Lecture des fichiers dans le répertoire
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Dim nomFichier As String
    
    For Each file In fileSystem.GetFolder(folderPath).Files
        nomFichier = file.Name
        Debug.Print nomFichier
        'Appliquer les traitements en fonction des fichiers
        Select Case True
            Case nomFichier = "LogSaisieHeures.log"
                Call LireLogSaisieHeures(file.path)
            Case nomFichier = "LogClientsApp.log"
                Call LireLogClientsApp(file.path)
            Case nomFichier = "LogMainApp.log"
                Call LireLogMainApp(file.path)
            Case Left(nomFichier, 11) = "LogMainApp." And Right(nomFichier, 4) = ".log"
                Call LireLogMainApp(file.path)
        End Select
    Next file
    
    'Libérer la mémoire
    Set file = Nothing
    Set fileDialog = Nothing
    Set fileSystem = Nothing
    
    MsgBox "Le traitement des fichiers LOG est terminé !", vbInformation
    
End Sub

Sub LireLogClientsApp(filePath As String)

    Application.StatusBar = "Traitement de '" & Fn_ExtraireNomFichier(filePath) & "'"
    
    'Ouvrir le fichier 'LogClientsApp.log'
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    'Détermine l'environnement (DEV/PROD) ?
    Dim env As String
    If Not InStr(filePath, wsdADMIN.Range("PATH_DATA_FILES") & Application.PathSeparator & gDATA_PATH) = 1 Then
        env = "PROD"
    Else
        env = "DEV"
    End If
    
    'Lire le fichier ligne par ligne et emmagasiner les champs dans un tableau
    Dim output() As Variant
    ReDim output(1 To 100000, 1 To 9)
    Dim ligne As Long
    Dim lineContent As String
    Dim lineNo As Long
    Dim duree As String
    Dim i As Long

    ligne = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineContent
        lineNo = lineNo + 1
        If lineNo Mod 250 = 0 Then
            Application.StatusBar = "Traitement de '" & Fn_ExtraireNomFichier(filePath) & "' - " & Format$(lineNo, "###,##0") & " lignes"
        End If
        If InStr(lineContent, " | ") <> 0 Then
            Dim Fields() As String
            Fields = Split(lineContent, " | ") 'Diviser la ligne en champs avec le délimiteur "|"
            'Insérer les données dans le tableau
            ligne = ligne + 1
            output(ligne, 1) = env
            output(ligne, 2) = CStr(Left$(Fields(0), 10))
            output(ligne, 3) = CStr(Right$(Fields(0), 11))
            output(ligne, 4) = Trim$(Fields(1))
            output(ligne, 5) = Trim$(Fields(2))
            output(ligne, 6) = Trim$(Fields(3))
            If InStr(Fields(3), " secondes'") <> 0 Then
                duree = Fn_ExtraireSecondesChaineLog(Fields(3))
                duree = Replace(duree, ".", ",")
                If duree <> 0 Then
                    output(ligne, 7) = CDbl(duree)
                Else
                    output(ligne, 7) = 0
                End If
                output(ligne, 6) = Trim$(Left$(output(ligne, 6), InStr(output(ligne, 6), " = ") - 1)) & " (S)"
            End If
            output(ligne, 8) = lineNo
            output(ligne, 9) = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        End If
    Loop

    'Réduit la taille du tableau output
    Call RedimensionnerTableau2D(output, ligne, UBound(output, 2))
    
    'Fermer le fichier
    Close #fileNum
    
    'Ajout du tableau à un classeur fermé
    Call AjouterTableauClasseurFerme(output, wsdADMIN.Range("PATH_DATA_FILES") & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & "GCF_Logs_Data.xlsb", "Log_Clients")
    
    'S'il s'agit du fichier DEV, on l'efface (on garde les fichiers logs de la PROD)
    Debug.Print env, filePath
    If env = "DEV" Then
        Kill filePath
    End If

    Application.StatusBar = False
    
End Sub

Sub LireLogMainApp(filePath As String)

    Application.StatusBar = "Traitement de '" & Fn_ExtraireNomFichier(filePath) & "'"
    
    'Ouvrir le fichier .Log
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    'Détermine l'environnement (DEV/PROD) ?
    Dim env As String
    If Not InStr(filePath, wsdADMIN.Range("PATH_DATA_FILES") & Application.PathSeparator & gDATA_PATH) = 1 Then
        env = "PROD"
    Else
        env = "DEV"
    End If
    
    'Lire le fichier ligne par ligne et emmagasiner les champs dans un tableau
    Dim output() As Variant
    ReDim output(1 To 200000, 1 To 10)
    Dim ligne As Long
    Dim lineNo As Long
    Dim lineContent As String
    Dim duree As String
    Dim i As Long

    ligne = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineContent
        If Not Trim(lineContent) = vbNullString Then
            lineNo = lineNo + 1
            If InStr(lineContent, " | ") <> 0 Then
                Dim Fields() As String
                Fields = Split(lineContent, " | ") 'Diviser la ligne en champs avec le délimiteur " | "
                'Insérer les données dans le tableau
                ligne = ligne + 1
                If ligne Mod 250 = 0 Then
                    Application.StatusBar = "Traitement de '" & Fn_ExtraireNomFichier(filePath) & "' - " & Format$(ligne, "###,##0") & " lignes"
                End If
                If UBound(Fields) = 5 Then
                    output(ligne, 1) = env
                    output(ligne, 2) = CStr(Left$(Fields(0), 10))
                    output(ligne, 3) = CStr(Right$(Fields(0), 11))
                    output(ligne, 4) = Trim$(Fields(1))
                    output(ligne, 5) = Trim$(Fields(2))
                    output(ligne, 6) = Trim$(Fields(3))
                    output(ligne, 7) = Trim$(Fields(4))
                    If InStr(Fields(5), " secondes") <> 0 Then
                        duree = Fn_ExtraireSecondesChaineLog(Fields(5))
                        duree = Replace(duree, ".", ",")
                        If duree <> 0 Then
                            output(ligne, 8) = CDbl(duree)
                        Else
                            output(ligne, 8) = 0
                        End If
                        output(ligne, 6) = Fields(3) & " (S)"
                    End If
                    output(ligne, 9) = lineNo
                    output(ligne, 10) = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
                End If
                If UBound(Fields) <= 4 Then
                    output(ligne, 1) = env
                    output(ligne, 2) = CStr(Left$(Fields(0), 10))
                    output(ligne, 3) = CStr(Right$(Fields(0), 11))
                    output(ligne, 4) = Trim$(Fields(1))
                    output(ligne, 5) = Trim$(Fields(2))
                    output(ligne, 6) = Trim$(Fields(3))
                    If UBound(Fields) = 4 Then
                        If InStr(Fields(4), " secondes") <> 0 Then
                            duree = Fn_ExtraireSecondesChaineLog(Fields(4))
                            duree = Replace(duree, ".", ",")
            '                    duree = Mid$(Fields(3), InStr(Fields(3), " *** = '") + 8)
            '                    duree = Left$(duree, InStr(duree, " ") - 1)
                            If duree <> 0 Then
                                output(ligne, 8) = CDbl(duree)
                            Else
                                output(ligne, 8) = 0
                            End If
    '                        output(ligne, 6) = Trim$(Left$(Fields(4), InStr(Fields(4), " = ") - 1)) & " (S)"
                        End If
                    End If
                    output(ligne, 9) = lineNo
                    output(ligne, 10) = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
                End If
            End If
        End If
    Loop

    'Réduit la taille du tableau output
    Call RedimensionnerTableau2D(output, ligne, UBound(output, 2))
    
    'Fermer le fichier
    Close #fileNum
    
    'Ajout du tableau au classeur des logs
    Call AjouterTableauClasseurFerme(output, wsdADMIN.Range("PATH_DATA_FILES") & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & "GCF_Logs_Data.xlsb", "Log_Application")
    
    'S'il s'agit du fichier DEV, on l'efface (on garde les fichiers logs de la PROD)
    If env = "DEV" Then
        Kill filePath
    End If
    
    Application.StatusBar = False
    
End Sub

Sub LireLogSaisieHeures(filePath As String)

    Application.StatusBar = "Traitement de '" & Fn_ExtraireNomFichier(filePath) & "'"
    
    'Ouvrir le fichier 'LogClientsApp.log'
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    'Détermine l'environnement (DEV/PROD) ?
    Dim env As String
    If Not InStr(filePath, wsdADMIN.Range("PATH_DATA_FILES") & Application.PathSeparator & gDATA_PATH) = 1 Then
        env = "PROD"
    Else
        env = "DEV"
    End If
    
    'Lire le fichier ligne par ligne et emmagasiner les champs dans un tableau
    Dim output() As Variant
    ReDim output(1 To 5000, 1 To 16)
    Dim ligne As Long
    Dim lineContent As String
    Dim lineNo As Long
    Dim duree As String
    Dim i As Long

    ligne = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineContent
        lineNo = lineNo + 1
        If lineNo Mod 250 = 0 Then
            Application.StatusBar = "Traitement de '" & Fn_ExtraireNomFichier(filePath) & "' - " & Format$(lineNo, "###,##0") & " lignes"
        End If
        If InStr(lineContent, " | ") <> 0 Then
            Dim Fields() As String
            Fields = Split(lineContent, " | ") 'Diviser la ligne en champs avec le délimiteur "|"
            'Insérer les données dans le tableau
            If UBound(Fields, 1) = 12 Then '2025-08-25 @ 20:19
                Fields(8) = Fields(8) & Fields(9)
                Fields(9) = Fields(10)
                Fields(10) = Fields(11)
            End If
            ligne = ligne + 1
            output(ligne, 1) = env
            output(ligne, 2) = CStr(Left$(Fields(0), 10))
            output(ligne, 3) = CStr(Right$(Fields(0), 11))
            output(ligne, 4) = Trim$(Fields(1))
            output(ligne, 5) = Trim$(Fields(2))
            Dim oper As String
            Dim tecID As Long
            oper = Trim$(Fields(3))
            tecID = Mid$(oper, 8, Len(oper) - 7)
            oper = Trim$(Left$(oper, 7))
            output(ligne, 6) = oper
            output(ligne, 7) = CStr(tecID)
            output(ligne, 8) = Fields(4)
            output(ligne, 9) = Fields(5)
            output(ligne, 10) = Fields(6)
            output(ligne, 11) = Fields(7)
            output(ligne, 12) = Fields(8)
            Dim hres As Double
            hres = CDbl(Replace(Fields(9), ".", ","))
            output(ligne, 13) = Round(hres, 2)
            output(ligne, 14) = Fields(10)
            output(ligne, 15) = lineNo
            output(ligne, 16) = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        End If
    Loop

    'Réduit la taille du tableau output
    Call RedimensionnerTableau2D(output, ligne, UBound(output, 2))
    
    'Fermer le fichier
    Close #fileNum
    
    'Ajout du tableau à un classeur fermé
    Call AjouterTableauClasseurFerme(output, wsdADMIN.Range("PATH_DATA_FILES") & Application.PathSeparator & gDATA_PATH & Application.PathSeparator & "GCF_Logs_Data.xlsb", "Log_Heures")
    
    'S'il s'agit du fichier DEV, on l'efface (on garde les fichiers logs de la PROD)
    Debug.Print env, filePath
    If env = "DEV" Then
        Kill filePath
    End If
    
    Application.StatusBar = False
    
End Sub

Sub AjouterTableauClasseurFerme(ByVal tableau As Variant, ByVal cheminFichier As String, ByVal feuilleNom As String)
    
    Dim wbSource As Workbook
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Ouvrir le classeur cible en arrière-plan
    Dim wbTarget As Workbook
    Set wbTarget = Workbooks.Open(cheminFichier)
    Dim wsTarget As Worksheet
    Set wsTarget = wbTarget.Sheets(feuilleNom)

    'Déterminer la première ligne vide dans la colonne A et définir le range
    Dim premiereLigneVide As Long
    premiereLigneVide = wsTarget.Cells(wsTarget.Rows.count, 1).End(xlUp).Row + 1
    Dim cible As Range
    Set cible = wsTarget.Cells(premiereLigneVide, 1)

    'Copier les données en une seule opération
    Application.EnableEvents = False
    cible.Resize(UBound(tableau, 1), UBound(tableau, 2)).Value = tableau
    Application.EnableEvents = True

    'Sauvegarder et fermer le fichier Target
    wbTarget.Close SaveChanges:=True

    'Libérer la mémoire
    Set cible = Nothing
    Set wbTarget = Nothing
    Set wsTarget = Nothing
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
