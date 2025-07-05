Attribute VB_Name = "modAuditVBA"
Option Explicit

Sub AnalyserToutesLesProcedures()

    Dim tableProc(1 To 1000, 1 To 8) As Variant '[Nom, Module, Type, Direct, Préfixé, Indirect, Object, NonConformite]
    Dim indexMax As Long
    Dim dictIndex As Object
    Set dictIndex = NouveauDict()

    Debug.Print "Début du traitement : analyse des procédures VBA"

    'Étape 1 : Construction du tableau
    Call BatirTableProcedures(dictIndex, tableProc, indexMax)

    'Étape 2 : Comptage des appels dans le code
    Call IncrementeAppelsViaIndex(dictIndex, tableProc, indexMax)

    'Étape 3 : Analyse des objets Excel (formes, boutons, etc.)
    Call ScannerObjetsExcelPourOnAction(dictIndex, tableProc)
    
    'Étape 4 : Export vers Excel trié et structuré
    Call ExporterToutesLesProcedures(tableProc, indexMax)

    Debug.Print "Traitement terminé (" & indexMax & " procédures analysées)"
    
    Application.ScreenUpdating = True
    Worksheets("ToutesLesProcedures").Activate

End Sub

Sub BatirTableProcedures(ByRef dictIndex As Object, ByRef tableProc() As Variant, ByRef index As Long)

    Debug.Print "   1. Construction de la liste des Procédures"
    
    Dim comp As Object, codeMod As Object
    Dim ligne As String, nomSub As String
    Dim i As Long, typeModule As String

    Set dictIndex = CreateObject("Scripting.Dictionary")
    index = 0

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Set codeMod = comp.codeModule

        Select Case comp.Type
            Case vbext_ct_StdModule: typeModule = "Module Standard"
            Case vbext_ct_ClassModule: typeModule = "Classe"
            Case vbext_ct_MSForm: typeModule = "UserForm"
            Case vbext_ct_Document: typeModule = "Feuille Excel"
            Case Else: typeModule = "Autre"
        End Select

        For i = 1 To codeMod.CountOfLines
            ligne = Trim(codeMod.Lines(i, 1))
            If Left(ligne, 1) = "'" Or InStr(ligne, "Function") > 0 Or InStr(ligne, "Sub ") = 0 Then GoTo NextLigne

            nomSub = codeMod.ProcOfLine(i, vbext_pk_Proc)

            If Not dictIndex.Exists(nomSub) Then
                index = index + 1
                tableProc(index, 1) = nomSub
                tableProc(index, 2) = comp.Name
                tableProc(index, 3) = typeModule
                tableProc(index, 4) = 0
                tableProc(index, 5) = 0
                tableProc(index, 6) = 0
                dictIndex.Add nomSub, index
            End If

NextLigne:
        Next i
    Next comp
    
End Sub

Sub IncrementeAppelsViaIndex(dictIndex As Object, tableProc() As Variant, indexMax As Long)

    Debug.Print "   2. Comptage des appels aux procédures (dans le code)"
    
    Dim comp As Object, lignes() As String, ligne As String
    Dim i As Long, nomProc As Variant, posEq As Long, valeur As String
    Dim pos1 As Long, pos2 As Long

    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Or comp.Type = vbext_ct_ClassModule Or _
           comp.Type = vbext_ct_MSForm Or comp.Type = vbext_ct_Document Then

            lignes = Split(comp.codeModule.Lines(1, comp.codeModule.CountOfLines), vbCrLf)

            For i = 0 To UBound(lignes)
                ligne = Trim(lignes(i))
                'Filtrage initial des lignes non pertinentes
                If ligne = "" Or Left(ligne, 1) = "'" Or _
                   LCase(Left(ligne, 12)) = "debug.print " Or _
                   LCase(Left(ligne, 7)) = "msgbox " Or _
                   LCase(Left(ligne, 4)) = "set " Or _
                   LCase(Left(ligne, 4)) = "sub " Or _
                   LCase(Left(ligne, 9)) = "function " Then
                    GoTo LigneSuivante
                End If

                For Each nomProc In dictIndex.keys
'                    If nomProc = "BoutonImprimer_CC" Then Stop

                    'Appels préfixés (Module.nomProc)
                    If InStr(ligne, "." & nomProc) > 0 Then
                        tableProc(dictIndex(nomProc), 5) = tableProc(dictIndex(nomProc), 5) + 1
                    End If

                    'Appels directs (nomProc)
                    If LCase(ligne) Like "*call *" Then
                        valeur = Trim(Split(LCase(ligne), "call")(1))
                        valeur = Replace(valeur, "()", "")
                        valeur = Split(valeur, " ")(0)
                        If valeur = LCase(nomProc) Then
                            tableProc(dictIndex(nomProc), 4) = tableProc(dictIndex(nomProc), 4) + 1
                        End If
                    End If
                    
'                    'Appels indirects via .OnAction
'                    If InStr(LCase(ligne), ".onaction") > 0 And InStr(ligne, "=") > 0 Then
'                        posEq = InStr(ligne, "=")
'                        valeur = Trim(Mid(ligne, posEq + 1)) 'Extrait le texte après le signe égal
'                        valeur = Replace(valeur, Chr(34), "")
'                        valeur = Replace(valeur, "'", "")
'                        valeur = Split(valeur, " ")(0) 'Garde le premier mot (souvent le nom de procédure)
'                        If LCase(valeur) = LCase(nomProc) Then
'                            tableProc(dictIndex(nomProc), 6) = tableProc(dictIndex(nomProc), 6) + 1
'                        End If
'                    End If

                    'Appels indirects dynamiques : Application.Run, Evaluate, Excel4Macro
                    If (InStr(LCase(ligne), "application.run") > 0 Or _
                        InStr(LCase(ligne), "evaluate(") > 0 Or _
                        InStr(LCase(ligne), "executeexcel4macro") > 0) Then

                        pos1 = InStr(ligne, """")
                        If pos1 > 0 Then
                            pos2 = InStr(pos1 + 1, ligne, """")
                            If pos2 > pos1 Then
                                valeur = Mid(ligne, pos1 + 1, pos2 - pos1 - 1)
                                valeur = Replace(valeur, "()", "")
                                valeur = Trim(Split(valeur, "!")(UBound(Split(valeur, "!")))) ' garde le nom après ! s'il est là
                                If LCase(valeur) = LCase(nomProc) Then
                                    tableProc(dictIndex(nomProc), 6) = tableProc(dictIndex(nomProc), 6) + 1
                                End If
                            End If
                        End If
                    End If
                Next nomProc

LigneSuivante:
            Next i
        End If
    Next comp
    
End Sub

Sub ScannerObjetsExcelPourOnAction(dictIndex As Object, tableProc() As Variant)

    Debug.Print "   3. Comptage des appels aux procédures (via Objets)"
    
    Dim ws As Worksheet
    Dim shp As Shape
    Dim obj As Object
    Dim nomMacro As String

    Dim idx As Long
    For Each ws In ThisWorkbook.Worksheets
        'Formes dessinées (Shapes)
        For Each shp In ws.Shapes
            If ws.Name = "FAC_Interrogation" Then Debug.Print shp.Name & " " & shp.OnAction: Stop
            If shp.OnAction <> "" Then
                nomMacro = shp.OnAction
                If nomMacro = "BoutonImprimer_CC" Then Stop
                If InStr(nomMacro, "!") > 0 Then
                    nomMacro = Split(nomMacro, "!")(1)
                End If
                If dictIndex.Exists(nomMacro) Then
                    idx = dictIndex(nomMacro)
                    tableProc(idx, 6) = tableProc(idx, 6) + 1
                    If tableProc(idx, 7) = "" Then
                        tableProc(idx, 7) = shp.Name & " (" & ws.Name & ")" 'nom de l’objet appelant
                    Else
                        tableProc(idx, 7) = tableProc(idx, 7) & vbCrLf & shp.Name & " (" & ws.Name & ")" 'nom de l’objet appelant
                    End If
                End If
            End If
        Next shp

        'Boutons de formulaire (si présents)
        For Each obj In ws.Buttons
            If obj.OnAction <> "" Then
                nomMacro = obj.OnAction
                If dictIndex.Exists(nomMacro) Then
                idx = dictIndex(nomMacro)
                    tableProc(idx, 6) = tableProc(idx, 6) + 1
                    If tableProc(idx, 7) = "" Then
                        tableProc(idx, 7) = obj.Name & " (" & ws.Name & ")" 'nom de l’objet appelant
                    Else
                        tableProc(idx, 7) = tableProc(idx, 7) & vbCrLf & obj.Name & " (" & ws.Name & ")" 'nom de l’objet appelant
                    End If
                End If
            End If
        Next obj
    Next ws
    
End Sub

Sub ExporterToutesLesProcedures(tableProc() As Variant, indexMax As Long)

    Debug.Print "   4. Exportation des résultats vers une feuille"

    Application.EnableEvents = False
    
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = Worksheets("ToutesLesProcedures")
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.count))
    ws.Name = "ToutesLesProcedures"

    Dim i As Long
    With ws
        'Entêtes
        .Cells(1, 1).Value = "Nom Procédure"
        .Cells(1, 2).Value = "Module"
        .Cells(1, 3).Value = "Type Module"
        .Cells(1, 4).Value = "Appels directs"
        .Cells(1, 5).Value = "Appels préfixés"
        .Cells(1, 6).Value = "Appels indirects"
        .Cells(1, 7).Value = "Total appels"
        .Cells(1, 8).Value = "Objet .OnAction"
        .Cells(1, 9).Value = "Non conformité"

        'Contenu
        For i = 1 To indexMax
            .Cells(i + 1, 1).Value = tableProc(i, 1)
            .Cells(i + 1, 2).Value = tableProc(i, 2)
            .Cells(i + 1, 3).Value = tableProc(i, 3)
            .Cells(i + 1, 4).Value = tableProc(i, 4) ' directs
            .Cells(i + 1, 5).Value = tableProc(i, 5) ' préfixés
            .Cells(i + 1, 6).Value = tableProc(i, 6) ' indirects
            .Cells(i + 1, 7).FormulaR1C1 = "=RC[-3]+RC[-2]+RC[-1]"
            .Cells(i + 1, 8).Value = tableProc(i, 7)
            If tableProc(i, 8) <> "" Then
                .Cells(i + 1, 9).Interior.Color = RGB(255, 230, 230) 'Rouge pâle
            End If
            .Cells(i + 1, 9).Value = tableProc(i, 8)
        Next i

        'Tri multicritère
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("A2:A" & indexMax + 1), Order:=xlAscending
            .SortFields.Add key:=ws.Range("B2:B" & indexMax + 1), Order:=xlAscending
            .SortFields.Add key:=ws.Range("C2:C" & indexMax + 1), Order:=xlAscending
            .SetRange ws.Range("A1:G" & indexMax + 1)
            .Header = xlYes
            .Apply
        End With
    End With
    
    Call DiagnostiquerConformite(ws, tableProc)
    
    'Mise en forme de la feuille
    With ws
        .Columns("A:I").AutoFit
        
        .Cells.VerticalAlignment = xlTop
    
        'Entêtes centrées, surbrillance bleue, texte blanc, gras, italique, taille réduite
        With .Range("A1:I1")
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(0, 102, 204) 'Bleu vif
            .Font.Color = vbWhite
            .Font.Bold = True
            .Font.Italic = True
            .Font.size = 9
        End With
    
        'Colonnes spécifiques centrées horizontalement
        .Columns("D").HorizontalAlignment = xlCenter
        .Columns("E").HorizontalAlignment = xlCenter
        .Columns("F").HorizontalAlignment = xlCenter
        .Columns("G").HorizontalAlignment = xlCenter
        .Columns("I").HorizontalAlignment = xlCenter
    
        'Filtre sur toutes les colonnes
        .Range("A1:I1").AutoFilter
    
        'Volet figé entre ligne 1 et 2
        .Range("B2").Select
        ActiveWindow.FreezePanes = True
    
        'Lignes zébrées gris pâle/blanc
        For i = 2 To indexMax + 1
            If i Mod 2 = 0 Then
                .Rows(i).Interior.Color = RGB(242, 242, 242) 'Gris très pâle
            Else
                .Rows(i).Interior.ColorIndex = xlNone
            End If
        Next i
        
        'Légende des non-conformités à la fin
        Dim lastRow As Long: lastRow = indexMax + 3
        .Cells(lastRow, 1).Value = "Légende des non-conformités :"
        .Cells(lastRow + 1, 1).Value = "R1 - Usage non autorisé de '_' sauf pour événements (_Click, _Change, etc)"
        .Cells(lastRow + 2, 1).Value = "R2 - Le nom contient un caractère accentué"
        .Cells(lastRow + 3, 1).Value = "R3 - Le nom ne commence pas par une majuscule"
        .Cells(lastRow + 4, 1).Value = "R4 - Le nom ne commence pas par un verbe d’action reconnu"
        .Cells(lastRow + 5, 1).Value = "R5 - La procédure n’est appelée nulle part"
    
        .Range(.Cells(lastRow, 1), .Cells(lastRow + 5, 1)).Font.size = 9
        .Range(.Cells(lastRow, 1), .Cells(lastRow + 5, 1)).Font.Italic = True
    End With
    
    ws.Activate
    ws.Select
    
    Application.EnableEvents = True
    
End Sub

Sub DiagnostiquerConformite(ws As Worksheet, tableProc() As Variant)

    Dim i As Long, nom As String, totalAppels As Long
    Dim diagnostics As String
    
    Dim dernLigneUtilisee As Long
    dernLigneUtilisee = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For i = 2 To dernLigneUtilisee
        nom = Trim(ws.Cells(i, 1).Value)
        
        diagnostics = ""

        'R1 - "_" non autorisé sauf si le nom SE TERMINE par un gestionnaire d’événement
        If InStr(nom, "_") > 0 And Not EstSuffixeEvenement(nom) Then
            diagnostics = diagnostics & "R1,"
        End If

        'R2 - Aucune accent dans les noms de procédures (choix personnel)
        If ContientAccent(nom) Then diagnostics = diagnostics & "R2,"

        'R3 - Doit commencer pas par une lettre majuscule
        If Left(nom, 1) <> UCase(Left(nom, 1)) Then diagnostics = diagnostics & "R3,"

        'R4 - Doit commencer par un verbe d’action
        If Not CommenceParVerbe(nom) Then diagnostics = diagnostics & "R4,"

        'R5 - Procédure n'est jamais appelé par l'application
        totalAppels = ws.Cells(i, 4).Value + ws.Cells(i, 5).Value + ws.Cells(i, 6).Value
        If totalAppels < 1 Then diagnostics = diagnostics & "R5,"

        If Right(diagnostics, 1) = "," Then diagnostics = Left(diagnostics, Len(diagnostics) - 1)
        
        If diagnostics <> "" Then
            ws.Cells(i, 9).Value = diagnostics
        End If
        
    Next i
    
End Sub

Function ContientAccent(texte As String) As Boolean

    Dim i As Long, code As Integer
    Const accents As String = "àâéèêîôùûäëïöüçÀÂÉÈÊÎÔÙÛÄËÏÖÜÇ"
    
    For i = 1 To Len(texte)
        If InStr(accents, Mid(texte, i, 1)) > 0 Then
            ContientAccent = True
            Exit Function
        End If
    Next i
    
    ContientAccent = False
    
End Function

Function EstSuffixeEvenement(nom As String) As Boolean

    Dim suffixes As Variant
    suffixes = Array("_AfterUpdate", "_BeforeClose", "_BeforeUpdate", "_Change", _
                     "_Click", "_Enter", "_Exit", "_SheetActivate", "_SheetChange")
                     
    Dim s As Variant
    Dim nbUnderscore As Long

    nbUnderscore = Len(nom) - Len(Replace(nom, "_", ""))

    'Si plus d'un underscore : rejet immédiat
    If nbUnderscore > 1 Then
        EstSuffixeEvenement = False
        Exit Function
    End If

    'Vérifie que le suffixe correspond à un événement reconnu
    For Each s In suffixes
        If LCase(Right(nom, Len(s))) = LCase(s) Then
            EstSuffixeEvenement = True
            Exit Function
        End If
    Next s

    EstSuffixeEvenement = False
    
End Function

Function CommenceParVerbe(nom As String) As Boolean

    Dim verbesAction As Variant
    verbesAction = Array("Activer", "Actualiser", "Afficher", "Ajouter", "Ajuster", _
                         "Aller", "Analyser", "Appliquer", "Assembler", "Batir", _
                         "Calculer", "Convertir", "Creer", "Effacer", "Executer", _
                         "Exporter", "Extraire", "Generer", "Importer", "Imprimer", _
                         "MettreAJour", "Nettoyer", "Preparer", "Redemmarer", "Remplir", _
                         "Reinitialiser", "Restaurer", "Sauvegarder", "Supprimer", _
                         "Traiter", "UserForm", "Valider", _
                         "Verifier", "Vider", "Workbook", "Worksheet")
    
    Dim v As Variant
    For Each v In verbesAction
        If LCase(Left(nom, Len(v))) = LCase(v) Then
            CommenceParVerbe = True
            Exit Function
        End If
    Next v
    CommenceParVerbe = False
    
End Function

Function NouveauDict() As Object

    Set NouveauDict = CreateObject("Scripting.Dictionary")
    
End Function


