Attribute VB_Name = "modAuditVBA"
Option Explicit

Sub zz_AnalyserToutesLesProcedures() '2025-08-05 @ 13:52

    Dim tableProc(1 To 1000, 1 To 8) As Variant '[Nom, Module, Type, Direct, Préfixé, Indirect, Object, NonConformite]
    Dim indexMax As Long
    Dim dictIndex As Object

    Debug.Print "Début du traitement : analyse des procédures VBA"

    'Étape 1 - Construction du tableau
    Call BatirTableProceduresEtFonctions(dictIndex, tableProc, indexMax)

    'Étape 2 - Comptage des appels dans le code
    Call IncrementerAppelsCodeDirect(dictIndex, tableProc, indexMax)

    'Étape 3 : Analyse des objets Excel (formes, boutons, etc.)
    Call IncrementerAppelsIndirect(dictIndex, tableProc)
    
    'Étape 4 : Export vers Excel trié et structuré
    Call ExporterResultatsFeuille(tableProc, indexMax)

    Debug.Print "Traitement terminé (" & indexMax & " procédures analysées)"
    
    Application.ScreenUpdating = True
    
    ActiveWorkbook.Worksheets("DocAuditVBA").Activate

End Sub

Sub BatirTableProceduresEtFonctions(ByRef dictIndex As Object, ByRef tableProc() As Variant, ByRef index As Long) '2025-07-07 @ 09:27

    Debug.Print "   1. Construction de la liste des Procédures"
    
    Dim comp As Object, codeMod As Object
    Dim ligne As String, nomSub As String
    Dim i As Long, typeModule As String

    Set dictIndex = CreateObject("Scripting.Dictionary")
    index = 0

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Set codeMod = comp.codeModule

        Select Case comp.Type
            Case vbext_ct_StdModule: typeModule = "3_Module Standard"
            Case vbext_ct_ClassModule: typeModule = "4_Classe"
            Case vbext_ct_MSForm: typeModule = "2_UserForm"
            Case vbext_ct_Document: typeModule = "1_Feuille Excel"
            Case vbext_ct_MSForm: typeModule = "2_UserForm"
            Case Else: typeModule = "z_Autre"
        End Select

        For i = 1 To codeMod.CountOfLines
            ligne = Trim(codeMod.Lines(i, 1))
            If Left(ligne, 1) = "'" Or InStr(ligne, "Function ") > 0 Or InStr(ligne, "Sub ") = 0 Then GoTo NextLigne

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

Sub IncrementerAppelsCodeDirect(dictIndex As Object, tableProc() As Variant, indexMax As Long) '2025-07-07 @ 09:27

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
                If ligne = vbNullString Or Left(ligne, 1) = "'" Or _
                   LCase(Left(ligne, 12)) = "debug.print " Or _
                   LCase(Left(ligne, 7)) = "msgbox " Or _
                   LCase(Left(ligne, 4)) = "set " Or _
                   LCase(Left(ligne, 4)) = "sub " Or _
                   LCase(Left(ligne, 9)) = "function " Then
                    GoTo LigneSuivante
                End If

'                If InStr(1, ligne, ".OnAction") Then Stop
                
                For Each nomProc In dictIndex.keys
'                    If nomProc = "EcrireInformationsConfigAuMenu" Then Stop
                    'Appels préfixés (Module.nomProc)
                    If InStr(ligne, "." & nomProc) > 0 Then
                        tableProc(dictIndex(nomProc), 5) = tableProc(dictIndex(nomProc), 5) + 1
                    End If

                    'Appels directs (nomProc)
                    If LCase(ligne) Like "*call *" Then
                        valeur = Trim(Split(LCase(ligne), "call")(1))
                        valeur = Replace(valeur, "()", vbNullString)
                        valeur = Split(valeur, " ")(0)
                        If InStr(valeur, "(") > 0 Then
                            valeur = Left(valeur, InStr(valeur, "(") - 1)
                        End If
                        If valeur = LCase(nomProc) Then
                            tableProc(dictIndex(nomProc), 4) = tableProc(dictIndex(nomProc), 4) + 1
                        End If
                    End If
                    
                    'Appels indirects dynamiques : Application.Run, Evaluate, Excel4Macro
                    If (InStr(LCase(ligne), "application.run") > 0 Or _
                        InStr(LCase(ligne), "evaluate(") > 0 Or _
                        InStr(LCase(ligne), "executeexcel4macro") > 0) Or _
                        InStr(LCase(ligne), ".onaction ") > 0 Then

                        pos1 = InStr(ligne, """")
                        If pos1 > 0 Then
                            pos2 = InStr(pos1 + 1, ligne, """")
                            If pos2 > pos1 Then
                                valeur = Mid(ligne, pos1 + 1, pos2 - pos1 - 1)
                                valeur = Replace(valeur, "()", vbNullString)
                                valeur = Trim(Split(valeur, "!")(UBound(Split(valeur, "!")))) 'garde le nom après ! s'il est là
                                If InStr(1, ligne, "'") Then
                                    valeur = Trim(Replace(valeur, "'", vbNullString))
                                End If
                                Debug.Print ligne
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

Sub IncrementerAppelsIndirect(dictIndex As Object, tableProc() As Variant) '2025-07-07 @ 09:27

    Debug.Print "   3. Comptage des appels aux procédures (via Objets)"
    
    Dim ws As Worksheet
    Dim shp As Shape
    Dim obj As Object
    Dim nomMacro As String

    Dim idx As Long
    For Each ws In ThisWorkbook.Worksheets
        'Formes dessinées (Shapes)
        For Each shp In ws.Shapes
            If shp.OnAction <> vbNullString Then
                nomMacro = shp.OnAction
                If InStr(nomMacro, "!") > 0 Then
                    nomMacro = Split(nomMacro, "!")(1)
                End If
                If dictIndex.Exists(nomMacro) Then
                    idx = dictIndex(nomMacro)
                    tableProc(idx, 6) = tableProc(idx, 6) + 1
                    If tableProc(idx, 7) = vbNullString Then
                        tableProc(idx, 7) = shp.Name & " (" & ws.Name & ")" 'nom de l’objet appelant
                    Else
                        tableProc(idx, 7) = tableProc(idx, 7) & vbCrLf & shp.Name & " (" & ws.Name & ")" 'nom de l’objet appelant
                    End If
                End If
            End If
        Next shp

        'Boutons de formulaire (si présents)
        For Each obj In ws.Buttons
            If obj.OnAction <> vbNullString Then
                nomMacro = obj.OnAction
                If dictIndex.Exists(nomMacro) Then
                idx = dictIndex(nomMacro)
                    tableProc(idx, 6) = tableProc(idx, 6) + 1
                    If tableProc(idx, 7) = vbNullString Then
                        tableProc(idx, 7) = obj.Name & " (" & ws.Name & ")"
                    Else
                        tableProc(idx, 7) = tableProc(idx, 7) & vbCrLf & obj.Name & " (" & ws.Name & ")"
                    End If
                End If
            End If
        Next obj
    Next ws
    
End Sub

Sub ExporterResultatsFeuille(tableProc() As Variant, indexMax As Long) '2025-07-07 @ 09:27

    Debug.Print "   4. Exportation des résultats vers une feuille"

    Application.EnableEvents = False
    
    'Utilisation de la feuille DocAuditVBA (permanente)
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("DocAuditVBA")
    ws.Cells.Clear 'Efface le contenu, mais garde les événements

    With ws
        'Légende interactive
        .Cells(1, 1).Value = "?? Double-cliquez sur un nom de procédure (colonne A) pour accéder directement au code VBA"
        .Cells(1, 1).Font.size = 10
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(0, 102, 204) 'Bleu
        .Cells(1, 1).Interior.Color = RGB(235, 247, 255) 'Bleu pâle
        .Cells(1, 1).HorizontalAlignment = xlCenter
        .Cells(1, 1).VerticalAlignment = xlCenter
        .Cells(1, 1).WrapText = True
        .Rows(1).RowHeight = 30
        'Entêtes
        .Cells(2, 1).Value = "Nom Procédure"
        .Cells(2, 2).Value = "Module"
        .Cells(2, 3).Value = "Type Module"
        .Cells(2, 4).Value = "Appels directs"
        .Cells(2, 5).Value = "Appels préfixés"
        .Cells(2, 6).Value = "Appels indirects"
        .Cells(2, 7).Value = "Total appels"
        .Cells(2, 8).Value = "Objet .OnAction"
        .Cells(2, 9).Value = "Non conformité"

        'Contenu
        Dim i As Long
        For i = 1 To indexMax
            .Cells(i + 2, 1).Value = tableProc(i, 1)
            .Cells(i + 2, 2).Value = tableProc(i, 2)
            .Cells(i + 2, 3).Value = tableProc(i, 3)
            .Cells(i + 2, 4).Value = tableProc(i, 4) 'directs
            .Cells(i + 2, 5).Value = tableProc(i, 5) 'préfixés
            .Cells(i + 2, 6).Value = tableProc(i, 6) 'indirects
            .Cells(i + 2, 7).FormulaR1C1 = "=RC[-3]+RC[-2]+RC[-1]"
            .Cells(i + 2, 8).Value = tableProc(i, 7)
            If tableProc(i, 8) <> vbNullString Then
                .Cells(i + 2, 9).Interior.Color = RGB(255, 230, 230) 'Rouge pâle
            End If
            .Cells(i + 2, 9).Value = tableProc(i, 8)
        Next i

        'Tri multicritère
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("C3:C" & indexMax + 1), Order:=xlAscending
            .SortFields.Add key:=ws.Range("B3:B" & indexMax + 1), Order:=xlAscending
            .SortFields.Add key:=ws.Range("A3:A" & indexMax + 1), Order:=xlAscending
            .SetRange ws.Range("A2:I" & indexMax + 2)
            .Header = xlYes
            .Apply
        End With
    End With
    
    Call DiagnostiquerConformite(ws, tableProc)
    
    'Mise en forme de la feuille
    With ws
        .Columns("A:I").AutoFit
        
        .Cells.VerticalAlignment = xlTop
    
        'Entêtes
        With .Range("A2:I2")
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(0, 102, 204) 'Bleu vif
            .Font.Color = vbWhite
            .Font.Bold = True
            .Font.Italic = True
            .Font.size = 9
        End With
    
        'Colonnes spécifiques
        .Columns("D").HorizontalAlignment = xlCenter
        .Columns("E").HorizontalAlignment = xlCenter
        .Columns("F").HorizontalAlignment = xlCenter
        .Columns("G").HorizontalAlignment = xlCenter
        .Columns("I").HorizontalAlignment = xlCenter
    
        'Filtre sur toutes les colonnes
        .Range("A2:I2").AutoFilter
    
        'Volet figé entre ligne 2 et 3
        .Activate
        .Range("B3").Select
        ActiveWindow.FreezePanes = True
    
        'Lignes zébrées bleu pâle/blanc
        For i = 3 To indexMax + 2
            If i Mod 2 = 0 Then
                .Rows(i).Interior.Color = RGB(220, 230, 241)
            Else
                .Rows(i).Interior.ColorIndex = xlNone
            End If
        Next i
        
        'Légende des non-conformités à la fin
        Dim lastRow As Long: lastRow = indexMax + 4
        .Cells(lastRow, 1).Value = "Légende des non-conformités :"
        .Cells(lastRow + 1, 1).Value = "R1 - Usage non autorisé de '_'sauf pour événements (_Click, _Change, etc)"
        .Cells(lastRow + 2, 1).Value = "R2 - Le nom contient un caractère accentué"
        .Cells(lastRow + 3, 1).Value = "R3 - Le nom ne commence pas par une majuscule"
        .Cells(lastRow + 4, 1).Value = "R4 - Le nom ne commence pas par un verbe d’action reconnu"
        .Cells(lastRow + 5, 1).Value = "R5 - La procédure n’est appelée nulle part"
    
        .Range(.Cells(lastRow, 1), .Cells(lastRow + 5, 1)).Font.size = 9
        .Range(.Cells(lastRow, 1), .Cells(lastRow + 5, 1)).Font.Italic = True
    End With
    
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    
        .LeftFooter = Format(Now, "yyyy-mm-dd") & " " & Format(Now, "hh:mm:ss")
        
        .CenterFooter = ws.Name
    
        .RightFooter = "Page &P de &N"
    
        'Marges serrées pour optimiser l’espace
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
    End With
    
    Application.EnableEvents = True
    
End Sub

Sub DiagnostiquerConformite(ws As Worksheet, tableProc() As Variant) '2025-07-07 @ 09:27

    Dim i As Long, nom As String, totalAppels As Long
    Dim diagnostics As String
    
    Dim dernLigneUtilisee As Long
    dernLigneUtilisee = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For i = 3 To dernLigneUtilisee
        nom = Trim(ws.Cells(i, 1).Value)
        
        diagnostics = vbNullString

        'R1 - "_" non autorisé sauf si le nom SE TERMINE par un gestionnaire d’événement
        If InStr(nom, "_") > 0 And Not EstSuffixeEvenement(nom) Then
            diagnostics = diagnostics & "R1,"
        End If

        'R2 - Aucune accent dans les noms de procédures (choix personnel)
        If ContientAccent(nom) Then diagnostics = diagnostics & "R2,"

        'R3 - Doit commencer pas par une lettre majuscule
        If Left(nom, 1) <> UCase(Left(nom, 1)) Then diagnostics = diagnostics & "R3,"

        'R4 - Doit commencer par un verbe d’action
        If Not ValiderCommenceParUnVerbe(nom) Then diagnostics = diagnostics & "R4,"

        'R5 - Procédure n'est jamais appelé par l'application
        totalAppels = ws.Cells(i, 4).Value + ws.Cells(i, 5).Value + ws.Cells(i, 6).Value
        If totalAppels < 1 Then diagnostics = diagnostics & "R5,"

        If Right(diagnostics, 1) = "," Then diagnostics = Left(diagnostics, Len(diagnostics) - 1)
        
        If diagnostics <> vbNullString Then
            ws.Cells(i, 9).Value = diagnostics
        End If
        
    Next i
    
End Sub

Function ContientAccent(texte As String) As Boolean '2025-07-07 @ 09:27

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

Function EstSuffixeEvenement(nom As String) As Boolean '2025-07-07 @ 09:27

    Dim suffixes As Variant
    suffixes = Array("_Activate", "_AfterUpdate", "_BeforeClose", "_BeforeDoubleClick", _
                     "_BeforeRightClick", "_BeforeUpdate", "_Change", "_Click", "_DblClick", _
                     "_Enter", "_Exit", "_Initialize", "_KeyDown", "_KeyUp", "_QueryClose", "_SelectionChange", _
                     "_SheetActivate", "_SheetChange", "_SheetDeactivate", "_SheetFollowHyperlink", _
                     "_SheetSelectionChange", "_Terminate")
                     
    Dim s As Variant
    Dim nbUnderscore As Long

    nbUnderscore = Len(nom) - Len(Replace(nom, "_", vbNullString))

    'Plus d'un underscore = rejet immédiat
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

Function ValiderCommenceParUnVerbe(nom As String) As Boolean '2025-07-07 @ 09:27

    Dim verbesAction As Variant
    verbesAction = Array("Acceder", "Activer", "Actualiser", "Additionner", "Afficher", "Ajouter", "Ajuster", _
                         "Aller", "Analyser", "Annuler", "Appeler", "Appliquer", "Arreter", "Assembler", "Batir", _
                         "Calculer", "Charger", "Cocher", "Compter", "Comparer", "Connecter", "Construire", "Convertir", _
                         "Copier", "Corriger", "Creer", "Decocher", "Demarrer", "Detecter", "Determiner", "Detruire", "Diagnostiquer", _
                         "Effacer", "Enregistrer", "Envoyer", "Executer", "Exporter", "Extraire", "Fermer", "Filtrer", "Fixer", _
                         "Fusionner", "Gerer", "Generer", "Identifier", "Importer", "Imprimer", "Incrementer", "Initialiser", "Inserer", _
                         "Lire", "Lister", "MettreAJour", "Nettoyer", "Noter", "Obtenir", "Planifier", "Positionner", "Preparer", _
                         "Rafraichir", "Rechercher", "Redefinir", "Redemmarer", "Reinitialiser", "Relancer", "Remplir", "Restaurer", _
                         "Retourner", "Saisir", "Sauvegarder", "Scanner", "Selectionner", "Supprimer", "Tester", "Traiter", "Transferer", _
                         "Trier", "UserForm", "Valider", "Verifier", "Vider", "Visualiser", "Workbook", "Worksheet", _
                         "btn", "chk", "cmb", "cmd", "ctrl", "shp", "txt", _
                         "DEB", "ENC", "FAC", "EJ", "GL", "REGUL", "TEC", "TEST")
    
    Dim v As Variant
    For Each v In verbesAction
        If LCase(Left(nom, Len(v))) = LCase(v) Then
            ValiderCommenceParUnVerbe = True
            Exit Function
        End If
    Next v
    ValiderCommenceParUnVerbe = False
    
End Function

Function AllerVersCode(nomModule As String, Optional nomProcedure As String = vbNullString) As Boolean '2025-07-07 @ 09:27

    On Error GoTo erreur

    Dim comp As VBComponent
    Dim cm As codeModule
    Dim cpane As CodePane
    Dim startLine As Long, numLines As Long

    'Recherche du module dans le projet VBA actif
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If Trim(comp.Name) = Trim(nomModule) Then Exit For
    Next

    If comp Is Nothing Then GoTo erreur

    comp.Activate

    'Si aucune procédure demandée, c'est terminé
    If nomProcedure = vbNullString Then
        AllerVersCode = True
        Exit Function
    End If

    'Recherche de la procédure dans le module
    Set cm = comp.codeModule
    startLine = cm.ProcStartLine(nomProcedure, vbext_pk_Proc)
    If startLine < 1 Then GoTo erreur

    numLines = cm.ProcCountLines(nomProcedure, vbext_pk_Proc)

    'Saut vers le bon CodePane
    For Each cpane In Application.VBE.CodePanes
        If cpane.codeModule Is cm Then
            With cpane
                .Window.Visible = True
                .SetSelection startLine, 1, startLine, 1
            End With

            AllerVersCode = True
            Exit Function
        End If
    Next

erreur:
    AllerVersCode = False
    
End Function

Sub BatirListeProcEtFoncDansDictionnaire() '2025-07-22 @ 12:01

    Dim comp As Object
    Dim codeMod As Object
    Dim ligne As String
    Dim nomSub As String
    Dim typeModule As String

    Dim dictProcFunc As Object
    Set dictProcFunc = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Set codeMod = comp.codeModule
        Select Case comp.Type
            Case vbext_ct_StdModule: typeModule = "3_Module Standard"
            Case vbext_ct_ClassModule: typeModule = "4_Classe"
            Case vbext_ct_MSForm: typeModule = "2_UserForm"
            Case vbext_ct_Document: typeModule = "1_Feuille Excel"
            Case vbext_ct_MSForm: typeModule = "2_UserForm"
            Case Else: typeModule = "z_Autre"
        End Select

        For i = 1 To codeMod.CountOfLines
            ligne = Trim(codeMod.Lines(i, 1))
            'Exclure les commentaires et les déclarations de fonction ou procédures
            If Left(ligne, 1) = "'" Or _
                InStr(ligne, "Function ") > 0 Or _
                InStr(ligne, "Sub") > 0 Then
                GoTo NextLigne
            End If

            nomSub = codeMod.ProcOfLine(i, vbext_pk_Proc)
            
            If nomSub <> vbNullString Then
                If Not dictProcFunc.Exists(nomSub) Then
                    dictProcFunc.Add nomSub, comp.Name
                End If
            End If

NextLigne:
        Next i
    Next comp
    
End Sub


