Attribute VB_Name = "modLog_Analysis"
Option Explicit

Sub Lire_Fichier_LogMainApp() '2025-01-13 @ 07:46

    'Utiliser une bo�te de dialogue Fichier
    Dim FileDialogBox As FileDialog
    Set FileDialogBox = Application.FileDialog(msoFileDialogFilePicker)

    'Configurer les filtres et afficher la bo�te de dialogue
    Dim nomCompletFichierLog As String
    With FileDialogBox
        .Title = "S�lectionnez le fichier 'LogMainApp' � analyser"
        .Filters.Clear 'Supprimer les filtres existants
        .Filters.Add "Fichiers log", "*.log"
        If .show = -1 Then
            nomCompletFichierLog = .selectedItems(1) 'R�cup�rer le chemin du fichier s�lectionn�
        Else
            MsgBox "Aucun fichier s�lectionn�.", vbExclamation
            Exit Sub
        End If
    End With

    'Est-ce le bon fichier de Log (format) ?
    If InStr(nomCompletFichierLog, "LogMainApp.log") = 0 Then
        MsgBox "Il ne s'agit pas du bon type de fichier Log", vbExclamation
        MsgBox "Traitement annul�", vbInformation
        Exit Sub
    End If
    
    'D�termine l'environnement (DEV/PROD) ?
    Dim env As String
    If Not InStr(nomCompletFichierLog, "C:\VBA\GC_FISCALIT�\DataFiles\") = 1 Then
        env = "PROD"
    Else
        env = "DEV"
    End If
    
    'Ouvrir le fichier s�lectionn�
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open nomCompletFichierLog For Input As #FileNumber

    'Lire le fichier ligne par ligne et emmagasiner les champs dans un tableau
    Dim output() As Variant
    ReDim output(1 To 50000, 1 To 9)
    Dim ligne As Long
    Dim LineContent As String
    Dim lineNo As Long
    Dim duree As String
    Dim i As Long

    ligne = 1
    Do While Not EOF(FileNumber)
        Line Input #FileNumber, LineContent
        lineNo = lineNo + 1
        If InStr(LineContent, " | ") <> 0 Then
            Dim Fields() As String
            Fields = Split(LineContent, " | ") 'Diviser la ligne en champs avec le d�limiteur "|"
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

    ' Fermer le fichier
    Close #FileNumber
    
    Call Array_2D_Resizer(output, ligne, UBound(output, 2))
    
    'Ajout du tableau � un classeur ferm�
    Call AjouterTableauClasseurFerme(output, "C:\VBA\GC_FISCALIT�\DataFiles\GCF_Logs_Data.xlsx", "Log_Application")

    'Lib�rer la m�moire
    Set FileDialogBox = Nothing

    MsgBox "Lecture du fichier LOG termin�e et donn�es ins�r�es dans la feuille.", vbInformation

End Sub

Sub AjouterTableauClasseurFerme(ByVal tableau As Variant, ByVal cheminFichier As String, ByVal feuilleNom As String)
    
    Dim cn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim lastUsedRow As Long
    Dim i As Long, j As Long
    
    'Est-ce bien un tableau ?
    If Not IsArray(tableau) Then
        MsgBox "Le param�tre 'tableau' doit �tre un tableau.", vbExclamation
        Exit Sub
    End If
    
    'Initialiser la connexion ADO
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = _
                    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & cheminFichier & ";" & _
                    "Extended Properties=""Excel 12.0 Xml;HDR=Yes"";"
    cn.Open
    
    'Obtenir la derni�re ligne utilis�e dans la feuille cible
    Set rs = cn.Execute("SELECT COUNT(*) AS NbLignes FROM [" & feuilleNom & "$]")
    lastUsedRow = rs.Fields("NbLignes").Value
    rs.Close
    
    'Boucle pour ins�rer les lignes du tableau dans le fichier Excel ferm�
    Dim valeur As Variant
    For i = LBound(tableau, 1) To UBound(tableau, 1)
        strSQL = "INSERT INTO [" & feuilleNom & "$] VALUES ("
        For j = LBound(tableau, 2) To UBound(tableau, 2)
            valeur = tableau(i, j)
            ' Ajouter des d�limiteurs dynamiquement en fonction du type de valeur
            If IsDate(valeur) Then
                strSQL = strSQL & "#" & Format(valeur, "yyyy-mm-dd hh:nn:ss") & "#, "
            ElseIf IsNumeric(valeur) And Not IsEmpty(valeur) Then
                strSQL = strSQL & Replace(valeur, ",", ".") & ", "
            ElseIf IsEmpty(valeur) Or IsNull(valeur) Then
                strSQL = strSQL & 0 & ", "
            Else
                strSQL = strSQL & "'" & Replace(valeur, "'", "''") & "', "
            End If
        Next j
        strSQL = Replace(strSQL, " 00:00:00#", "#")
        strSQL = Left(strSQL, Len(strSQL) - 2) & ")" 'Supprime la derni�re virgule
        cn.Execute strSQL
    Next i
    
    'Fermer la connexion
    cn.Close
    Set cn = Nothing
    
    MsgBox "Ajout termin� avec succ�s apr�s la ligne " & lastUsedRow & " !"
    
End Sub


