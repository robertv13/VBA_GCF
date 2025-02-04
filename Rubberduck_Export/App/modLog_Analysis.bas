Attribute VB_Name = "modLog_Analysis"
Option Explicit

Sub Main_OuvrirRepertoireEtTraiterFichiers()

    'Initialisation du FileDialog pour s�lectionner un r�pertoire
    Dim fileDialog As fileDialog
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    fileDialog.Title = "S�lectionnez un r�pertoire � traiter"
    
    'Un r�pertoire a-t-il �t� s�lectionn� ?
    Dim folderPath As String
    If fileDialog.show = -1 Then
        folderPath = fileDialog.selectedItems(1)
    Else
        MsgBox "Aucun r�pertoire s�lectionn�.", vbExclamation
        Exit Sub
    End If
    
    'V�rification de l'existence du r�pertoire
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "R�pertoire invalide.", vbCritical
        Exit Sub
    End If
    
    'Lecture des fichiers dans le r�pertoire
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    For Each file In fileSystem.GetFolder(folderPath).Files
        'Appliquer les traitements en fonction des fichiers
        Select Case file.Name
            Case "LogClientsApp.log"
                Call Lire_LogClientsApp(file.path)
            Case "LogMainApp.log"
                Call Lire_LogMainApp(file.path)
            Case "LogSaisieHeures.log"
                Call Lire_LogSaisieHeures(file.path)
        End Select
    Next file
    
    'Lib�rer la m�moire
    Set file = Nothing
    Set fileDialog = Nothing
    Set fileSystem = Nothing
    
    MsgBox "Le traitement des fichiers LOG est termin� !", vbInformation
    
End Sub

Sub Lire_LogClientsApp(filePath As String)

    Application.StatusBar = "Traitement de '" & ExtraireNomFichier(filePath) & "' - 0 ligne"
    
    'Ouvrir le fichier 'LogClientsApp.log'
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    'D�termine l'environnement (DEV/PROD) ?
    Dim env As String
    If Not InStr(filePath, "C:\VBA\GC_FISCALIT�\DataFiles\") = 1 Then
        env = "PROD"
    Else
        env = "DEV"
    End If
    
    'Lire le fichier ligne par ligne et emmagasiner les champs dans un tableau
    Dim output() As Variant
    ReDim output(1 To 25000, 1 To 9)
    Dim ligne As Long
    Dim lineContent As String
    Dim lineNo As Long
    Dim duree As String
    Dim i As Long

    ligne = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineContent
        lineNo = lineNo + 1
        If lineNo Mod 25 = 0 Then
            Application.StatusBar = "Traitement de '" & ExtraireNomFichier(filePath) & "' - " & Format$(lineNo, "###,##0") & " lignes"
        End If
        If InStr(lineContent, " | ") <> 0 Then
            Dim Fields() As String
            Fields = Split(lineContent, " | ") 'Diviser la ligne en champs avec le d�limiteur "|"
            'Ins�rer les donn�es dans le tableau
            ligne = ligne + 1
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
                If duree <> 0 Then
                    output(ligne, 7) = CDbl(duree)
                Else
                    output(ligne, 7) = 0
                End If
                output(ligne, 6) = Trim(Left(Fields(3), InStr(Fields(3), " = ") - 1)) & " (S)"
            End If
            output(ligne, 8) = lineNo
            output(ligne, 9) = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        End If
    Loop

    'R�duit la taille du tableau output
    Call Array_2D_Resizer(output, ligne, UBound(output, 2))
    
    'Fermer le fichier
    Close #fileNum
    
    'Ajout du tableau � un classeur ferm�
    Call AjouterTableauClasseurFerme(output, "C:\VBA\GC_FISCALIT�\DataFiles\GCF_Logs_Data.xlsb", "Log_Clients")
    
    'S'il s'agit du fichier DEV, on l'efface (on garde les fichiers logs de la PROD)
    Debug.Print env, filePath
    If env = "DEV" Then
        Kill filePath
    End If

    Application.StatusBar = ""
    
    'Afficher le nombre de lignes ajout�es au fichier LOG
    MsgBox "Le fichier '" & ExtraireNomFichier(filePath) & "' a ajout� " & Format$(UBound(output, 1), "###,##0") & " lignes au fichier cumulatif", vbInformation
    
End Sub

Sub Lire_LogMainApp(filePath As String)

    Application.StatusBar = "Traitement de '" & ExtraireNomFichier(filePath) & "'"
    
    'Ouvrir le fichier .Log
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    'D�termine l'environnement (DEV/PROD) ?
    Dim env As String
    If Not InStr(filePath, "C:\VBA\GC_FISCALIT�\DataFiles\") = 1 Then
        env = "PROD"
    Else
        env = "DEV"
    End If
    
    'Lire le fichier ligne par ligne et emmagasiner les champs dans un tableau
    Dim output() As Variant
    ReDim output(1 To 100000, 1 To 10)
    Dim ligne As Long
    Dim lineNo As Long
    Dim lineContent As String
    Dim duree As String
    Dim i As Long

    ligne = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineContent
        lineNo = lineNo + 1
        If InStr(lineContent, " | ") <> 0 Then
            Dim Fields() As String
            Fields = Split(lineContent, " | ") 'Diviser la ligne en champs avec le d�limiteur " | "
            'Ins�rer les donn�es dans le tableau
            ligne = ligne + 1
            If ligne Mod 250 = 0 Then
                Application.StatusBar = "Traitement de '" & ExtraireNomFichier(filePath) & "' - " & Format$(ligne, "###,##0") & " lignes"
            End If
            If UBound(Fields) = 5 Then
                output(ligne, 1) = env
                output(ligne, 2) = CStr(Left(Fields(0), 10))
                output(ligne, 3) = CStr(Right(Fields(0), 11))
                output(ligne, 4) = Trim(Fields(1))
                output(ligne, 5) = Trim(Fields(2))
                output(ligne, 6) = Trim(Fields(3))
                output(ligne, 7) = Trim(Fields(4))
                If InStr(Fields(5), " secondes") <> 0 Then
                    duree = ExtraireSecondes(Fields(5))
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
            If UBound(Fields) = 4 Then
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
                    If duree <> 0 Then
                        output(ligne, 7) = CDbl(duree)
                    Else
                        output(ligne, 7) = 0
                    End If
                    output(ligne, 6) = Trim(Left(Fields(3), InStr(Fields(3), " = ") - 1)) & " (S)"
                End If
                output(ligne, 8) = lineNo
                output(ligne, 9) = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            End If
        End If
    Loop

    Application.StatusBar = "Traitement de '" & ExtraireNomFichier(filePath) & "' - " & Format$(lineNo, "###,##0") & " lignes"

    'R�duit la taille du tableau output
    Call Array_2D_Resizer(output, ligne, UBound(output, 2))
    
    'Fermer le fichier
    Close #fileNum
    
    'Ajout du tableau au classeur des logs
    Call AjouterTableauClasseurFerme(output, "C:\VBA\GC_FISCALIT�\DataFiles\GCF_Logs_Data.xlsb", "Log_Application")
    
    'S'il s'agit du fichier DEV, on l'efface (on garde les fichiers logs de la PROD)
    If env = "DEV" Then
        Kill filePath
    End If
    
    Application.StatusBar = ""
    
    'Afficher le nombre de lignes ajout�es au fichier LOG
    MsgBox "Le fichier '" & ExtraireNomFichier(filePath) & "' a ajout� " & Format$(UBound(output, 1), "###,##0") & " lignes au fichier cumulatif", vbInformation
    
End Sub

Sub Lire_LogSaisieHeures(filePath As String)

    Application.StatusBar = "Traitement de '" & ExtraireNomFichier(filePath) & "'"
    
    'Ouvrir le fichier 'LogClientsApp.log'
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    'D�termine l'environnement (DEV/PROD) ?
    Dim env As String
    If Not InStr(filePath, "C:\VBA\GC_FISCALIT�\DataFiles\") = 1 Then
        env = "PROD"
    Else
        env = "DEV"
    End If
    
    'Lire le fichier ligne par ligne et emmagasiner les champs dans un tableau
    Dim output() As Variant
    ReDim output(1 To 2500, 1 To 16)
    Dim ligne As Long
    Dim lineContent As String
    Dim lineNo As Long
    Dim duree As String
    Dim i As Long

    ligne = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineContent
        lineNo = lineNo + 1
        If lineNo Mod 25 = 0 Then
            Application.StatusBar = "Traitement de '" & ExtraireNomFichier(filePath) & "' - " & Format$(lineNo, "###,##0") & " lignes"
        End If
        If InStr(lineContent, " | ") <> 0 Then
            Dim Fields() As String
            Fields = Split(lineContent, " | ") 'Diviser la ligne en champs avec le d�limiteur "|"
            'Ins�rer les donn�es dans le tableau
            ligne = ligne + 1
            output(ligne, 1) = env
            output(ligne, 2) = CStr(Left(Fields(0), 10))
            output(ligne, 3) = CStr(Right(Fields(0), 11))
            output(ligne, 4) = Trim(Fields(1))
            output(ligne, 5) = Trim(Fields(2))
            Dim oper As String
            Dim tecID As Long
            oper = Trim(Fields(3))
            tecID = Mid(oper, 8, Len(oper) - 7)
            oper = Trim(Left(oper, 7))
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

    'R�duit la taille du tableau output
    Call Array_2D_Resizer(output, ligne, UBound(output, 2))
    
    'Fermer le fichier
    Close #fileNum
    
    'Ajout du tableau � un classeur ferm�
    Call AjouterTableauClasseurFerme(output, "C:\VBA\GC_FISCALIT�\DataFiles\GCF_Logs_Data.xlsb", "Log_Heures")
    
    'S'il s'agit du fichier DEV, on l'efface (on garde les fichiers logs de la PROD)
    Debug.Print env, filePath
    If env = "DEV" Then
        Kill filePath
    End If
    
    Application.StatusBar = ""
    
    'Afficher le nombre de lignes ajout�es au fichier LOG
    MsgBox "Le fichier '" & ExtraireNomFichier(filePath) & "' a ajout� " & Format$(UBound(output, 1), "###,##0") & " lignes au fichier cumulatif", vbInformation
    
End Sub

Sub AjouterTableauClasseurFerme(ByVal tableau As Variant, ByVal cheminFichier As String, ByVal feuilleNom As String)
    
    Dim wbSource As Workbook
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Ouvrir le classeur cible en arri�re-plan
    Dim wbTarget As Workbook
    Set wbTarget = Workbooks.Open(cheminFichier)
    Dim wsTarget As Worksheet
    Set wsTarget = wbTarget.Sheets(feuilleNom)

    'D�terminer la premi�re ligne vide dans la colonne A et d�finir le range
    Dim premiereLigneVide As Long
    premiereLigneVide = wsTarget.Cells(wsTarget.Rows.count, 1).End(xlUp).row + 1
    Dim cible As Range
    Set cible = wsTarget.Cells(premiereLigneVide, 1)

    'Copier les donn�es en une seule op�ration
    Application.EnableEvents = False
    cible.Resize(UBound(tableau, 1), UBound(tableau, 2)).Value = tableau
    Application.EnableEvents = True

    'Sauvegarder et fermer le fichier Target
    wbTarget.Close SaveChanges:=True

    'Lib�rer la m�moire
    Set cible = Nothing
    Set wbTarget = Nothing
    Set wsTarget = Nothing
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

'    MsgBox "Donn�es copi�es avec succ�s dans " & vbNewLine & vbNewLine & _
'                    "'" & cheminFichier & "'", vbInformation
'
'    Dim cn As Object
'    Dim rs As Object
'    Dim strSQL As String
'    Dim lastUsedRow As Long
'    Dim i As Long, j As Long
'
'    'Est-ce bien un tableau ?
'    If Not IsArray(tableau) Then
'        MsgBox "Le param�tre 'tableau' doit �tre un tableau.", vbExclamation
'        Exit Sub
'    End If
'
'    'Initialiser la connexion ADO
'    Set cn = CreateObject("ADODB.Connection")
'    cn.ConnectionString = _
'                    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'                    "Data Source=" & cheminFichier & ";" & _
'                    "Extended Properties=""Excel 12.0 Xml;HDR=Yes"";"
'    cn.Open
'
'    'Obtenir la derni�re ligne utilis�e dans la feuille cible
'    Set rs = cn.Execute("SELECT COUNT(*) AS NbLignes FROM [" & feuilleNom & "$]")
'    lastUsedRow = rs.Fields("NbLignes").Value
'    rs.Close
'
'    'Boucle pour ins�rer les lignes du tableau dans le fichier Excel ferm�
'    Dim valeur As Variant
'    For i = LBound(tableau, 1) To UBound(tableau, 1)
'        strSQL = "INSERT INTO [" & feuilleNom & "$] VALUES ("
'        For j = LBound(tableau, 2) To UBound(tableau, 2)
'            'Nettoyage de la valeur
'            If Not IsEmpty(valeur) Then valeur = Trim(valeur)
'
'            'Tronque les donn�es qui seraient trop longues
'            valeur = tableau(i, j)
'            If Len(valeur) > 197 Then
'                valeur = Left(valeur, 197) & "..."
'            End If
'
'            'D�terminer dynamiquement le type de valeur
'            If IsEmpty(valeur) Or IsNull(valeur) Then
'                'Valeur vide ou nulle, ins�rer une valeur par d�faut
'                strSQL = strSQL & "0, "
'            ElseIf IsDate(valeur) Then
'                'Date : Format SQL compatible
'                strSQL = strSQL & "#" & Format(valeur, "yyyy-mm-dd hh:nn:ss") & "#, "
'            ElseIf IsNumeric(valeur) Then
'                If InStr(1, CStr(valeur), ".") > 0 Or InStr(1, CStr(valeur), ",") > 0 Then
'                    'La valeur contient un s�parateur d�cimal
'                    strSQL = strSQL & Replace(CDbl(valeur), ",", ".") & ", "
''                    strSQL = strSQL & Replace(Format(CDbl(valeur), "0.00"), ",", ".") & ", "
'                Else
'                    'La valeur est enti�re
'                    strSQL = strSQL & CLng(valeur) & ", "
'                End If
'            Else
'                'Texte : Prot�ger les apostrophes
'                strSQL = strSQL & "'" & Replace(valeur, "'", "''") & "', "
'            End If
'        Next j
'        strSQL = Replace(strSQL, " 00:00:00#", "#")
'        strSQL = Left(strSQL, Len(strSQL) - 2) & ")" 'Supprime la derni�re virgule
'        cn.Execute strSQL
'    Next i
'
'    'Fermer la connexion
'    cn.Close
'
'    'Lib�rer la m�moire
'    Set cn = Nothing
    
End Sub
