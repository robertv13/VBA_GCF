Attribute VB_Name = "modTBD"
Option Explicit

Sub AdditionnerSoldes(r1 As Range, r2 As Range, comptes As String)

    If comptes = vbNullString Then
        Exit Sub
    End If
    
    Dim compte() As String
    compte = Split(comptes, "^")
    
    Dim i As Integer
    For i = 0 To UBound(compte, 1) - 1
        r1.Value = r1.Value + ChercherSoldes(compte(i), 1)
    Next i

    r1.Value = Round(r1.Value, 0)
    
End Sub

'Ajustements à la feuille DB_Clients (Ajout du contactdans le nom du client)
Sub AjouterContactDansNomClient()

    'Declare and open the closed workbook
    Dim wb As Workbook: Set wb = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx")

    'Define the worksheet you want to work with
    Dim ws As Worksheet: Set ws = wb.Worksheets("Clients")
    
    'Find the last used row with data in column A
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    'Loop through each row starting from row 2 (headers are 1 row)
    Dim client As String, clientID As String, contactFacturation As String
    Dim posOpenSquareBracket As Integer, posCloseSquareBracket As Integer
'    Dim numberOpenSquareBracket As Integer, numberCloseSquareBracket As Integer
    Dim i As Long
    For i = 2 To lastUsedRow
        'Load data into variables
        client = ws.Cells(i, fClntFMClientNom).Value
        clientID = ws.Cells(i, fClntFMClientID).Value
        contactFacturation = Trim$(ws.Cells(i, fClntFMContactFacturation).Value)
        
        'Process the data and make adjustments if necessary
        posOpenSquareBracket = InStr(client, "[")
        posCloseSquareBracket = InStr(client, "]")
        
        If posOpenSquareBracket = 0 And posCloseSquareBracket = 0 Then
            If contactFacturation <> vbNullString And InStr(client, contactFacturation) = 0 Then
                client = Trim$(client) & " [" & contactFacturation & "]"
                ws.Cells(i, 1).Value = client
                Debug.Print "#065 - " & i & " - " & client
            End If
        End If
        
    Next i
    
    wb.Save
    
    'Libérer la mémoire
    Set wb = Nothing
    Set ws = Nothing
    
    MsgBox "Le traitement est complété sur " & i - 1 & " lignes"
    
End Sub

'Ajustements à la feuille DB_Clients (*) ---> [*]
Sub AjusterNomClientBD()

    'Declare and open the closed workbook
    Dim wb As Workbook: Set wb = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx")

    'Define the worksheet you want to work with
    Dim ws As Worksheet: Set ws = wb.Worksheets("Clients")
    
    'Find the last used row with data in column A
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    'Loop through each row starting from row 2 (headers are 1 row)
    Dim client As String, clientID As String, contactFacturation As String
    Dim posOpenParenthesis As Integer, posCloseParenthesis As Integer
    Dim numberOpenParenthesis As Integer, numberCloseParenthesis As Integer
    Dim i As Long
    For i = 2 To lastUsedRow
        'Load data into variables
        client = ws.Cells(i, fClntFMClientNom).Value
        clientID = ws.Cells(i, fClntFMClientID).Value
        contactFacturation = ws.Cells(i, fClntFMContactFacturation).Value
        
        'Process the data and make adjustments if necessary
        posOpenParenthesis = InStr(client, "(")
        posCloseParenthesis = InStr(client, ")")
        numberOpenParenthesis = Fn_Count_Char_Occurrences(client, "(")
        numberCloseParenthesis = Fn_Count_Char_Occurrences(client, ")")
        
        If numberOpenParenthesis = 1 And numberCloseParenthesis = 1 Then
            If posCloseParenthesis > posOpenParenthesis + 5 Then
                client = Replace(client, "(", "[")
                client = Replace(client, ")", "]")
                ws.Cells(i, 1).Value = client
                Debug.Print "#064 - " & i & " - " & client
            End If
        End If
        
    Next i
    
    wb.Save
    
    'Libérer la mémoire
    Set wb = Nothing
    Set ws = Nothing
    
    MsgBox "Le traitement est complété sur " & i - 1 & " lignes"
    
End Sub

Sub AnalyserImagesEnteteFactureExcel() '2025-05-27 @ 14:40

    Dim dossier As String, fichier As String
    Dim wb As Workbook, ws As Worksheet
    Dim img As Shape
    Dim largeurOrig As Double, hauteurOrig As Double
    Dim largeurActuelle As Double, hauteurActuelle As Double
    Dim cheminComplet As String
    Dim nomImageCible As String

    'Demande à l'utilisateur de choisir un dossier
    With Application.fileDialog(msoFileDialogFolderPicker)
        .Title = "Choisissez un dossier contenant les fichiers Excel"
        If .show <> -1 Then Exit Sub 'Annuler
        dossier = .SelectedItems(1)
    End With

    'Nom exact de l'image à trouver (ou utiliser un critère partiel)
    nomImageCible = "Image 1" '? Modifier si nécessaire

    'Recherche tous les fichiers .xlsx dans le dossier
    Dim dateSeuilMinimum As Date
    dateSeuilMinimum = DateSerial(2024, 8, 1)
    fichier = Dir(dossier & "\*.xlsx")

    Do While fichier <> vbNullString
        cheminComplet = dossier & "\" & fichier
        If FileDateTime(cheminComplet) < dateSeuilMinimum Then
            fichier = Dir
            GoTo SkipFile
        End If
        Set wb = Workbooks.Open(cheminComplet, ReadOnly:=True)

        On Error Resume Next
        Set ws = wb.Worksheets(wb.Worksheets.count)
        If ws.Name = "Activités" Then
            GoTo SkipFile
        End If
        On Error GoTo 0

        If Not ws Is Nothing Then
            For Each img In ws.Shapes
                If img.Type = msoPicture Then
                    If img.Name = nomImageCible Then
                        largeurActuelle = img.Width
                        hauteurActuelle = img.Height

                        'Lire la taille originale estimée
                        Call LireTailleOriginaleImage(img, largeurOrig, hauteurOrig)

                        Debug.Print "Fichier : " & fichier
                        Debug.Print "  Image : " & img.Name
                        Debug.Print "  Taille actuelle : " & largeurActuelle & " x " & hauteurActuelle
                        Debug.Print "  Taille originale : " & largeurOrig & " x " & hauteurOrig
                        Debug.Print String(40, "-")
                    End If
                End If
            Next img
        End If

        wb.Close SaveChanges:=False
        fichier = Dir
SkipFile:
    Loop

    MsgBox "Analyse terminée."
    
End Sub


