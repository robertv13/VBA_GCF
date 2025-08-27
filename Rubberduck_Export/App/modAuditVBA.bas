Attribute VB_Name = "modAuditVBA"
Option Explicit

Sub shpAuditVBAProcedures()

    Call AnalyserTousLesNomsDeProcedures

End Sub

Sub AnalyserTousLesNomsDeProcedures() '2025-08-05 @ 13:52

    Dim tableProc(1 To 750, 1 To 9) As Variant '[Nom, TypeProc, Module, TypeMod, Direct, Préfixé, Indirect, Object, NonConformite]
    Dim indexMax As Long
    Dim dictIndex As Object

    Application.ScreenUpdating = False
    
    Debug.Print "Début du traitement : analyse des procédures VBA"

    'Étape 1 - Construction du tableau
    Call BatirTableProceduresEtFonctions(dictIndex, tableProc, indexMax)

    'Étape 2 - Comptage des appels directs aux procédures dans le code
    Call IncrementerAppelsCodeDirect(dictIndex, tableProc, indexMax)

    'Étape 3 - Comptage des appels indirects aux procédures dans le code
    Call IncrementerAppelsIndirect(dictIndex, tableProc)
    
    'Étape 4 - Comptage des appels aux fonctions dans le code
    Call IncrementerAppelsFonctionsCodeDirect(dictIndex, tableProc, indexMax)
    
    'Étape 5 : Export vers Excel trié et structuré
    Call ExporterResultatsFeuille(tableProc, indexMax)

    Debug.Print "Traitement terminé (" & indexMax & " procédures analysées)"
    
    Application.ScreenUpdating = True

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
            Case Else: typeModule = "z_Autre"
        End Select

        For i = 1 To codeMod.CountOfLines
            ligne = Trim(codeMod.Lines(i, 1))
            If Not ligne Like "Sub *" And Not ligne Like "Function *" Then GoTo NextLigne

            nomSub = codeMod.ProcOfLine(i, vbext_pk_Proc)

            If Not dictIndex.Exists(nomSub) Then
                index = index + 1
                tableProc(index, 1) = nomSub
                tableProc(index, 2) = IIf(ligne Like "Sub *", "Sub", "Function")
                tableProc(index, 3) = comp.Name
                tableProc(index, 4) = typeModule
                tableProc(index, 5) = 0
                tableProc(index, 6) = 0
                tableProc(index, 7) = 0
                dictIndex.Add nomSub, index
                If nomSub = "Fn_AccesServeur" Then
                    Debug.Print nomSub, index
                End If
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
                   LCase(Left(ligne, 4)) = "sub " Or _
                   LCase(Left(ligne, 9)) = "function " Then
                    GoTo LigneSuivante
                End If

                For Each nomProc In dictIndex.keys
                    'Appels préfixés (Module.nomProc)
                    If InStr(ligne, "." & nomProc) > 0 Then
                        tableProc(dictIndex(nomProc), 5) = tableProc(dictIndex(nomProc), 6) + 1
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
                            tableProc(dictIndex(nomProc), 5) = tableProc(dictIndex(nomProc), 5) + 1
                        End If
                    End If
                    
                    'Appels indirects dynamiques : Application.Run, Evaluate, Excel4Macro & OnAction
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
                                If nomProc = "FermerApplicationInactive" And InStr(ligne, "FermerApplicationInactive") Then Stop
                                If LCase(valeur) = LCase(nomProc) Then
                                    tableProc(dictIndex(nomProc), 7) = tableProc(dictIndex(nomProc), 7) + 1
                                End If
                            End If
                        End If
                    End If
                    
                    'Appels indirects dynamiques : OnTime
                    If InStr(LCase(ligne), ".ontime") > 0 Then

                        pos1 = InStr(ligne, """")
                        If pos1 > 0 Then
                            pos2 = InStr(pos1 + 1, ligne, """")
                            If pos2 > pos1 Then
                                valeur = Mid(ligne, pos1 + 1, pos2 - pos1 - 1)
                                valeur = Replace(valeur, "()", vbNullString)
                                valeur = Trim(Split(valeur, "!")(UBound(Split(valeur, "!")))) 'garde le nom après ! s'il est là
                                If LCase(valeur) = LCase(nomProc) Then
                                    tableProc(dictIndex(nomProc), 7) = tableProc(dictIndex(nomProc), 7) + 1
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
            On Error Resume Next
            nomMacro = shp.OnAction
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo NextShape
            End If
            On Error GoTo 0
            If shp.OnAction <> vbNullString Then
                nomMacro = shp.OnAction
                If InStr(nomMacro, "!") > 0 Then
                    nomMacro = Split(nomMacro, "!")(1)
                End If
                If dictIndex.Exists(nomMacro) Then
                    idx = dictIndex(nomMacro)
                    tableProc(idx, 7) = tableProc(idx, 7) + 1
                    If tableProc(idx, 8) = vbNullString Then
                        tableProc(idx, 8) = shp.Name & " (" & ws.Name & ")" 'nom de l’objet appelant
                    Else
                        tableProc(idx, 8) = tableProc(idx, 8) & vbCrLf & shp.Name & " (" & ws.Name & ")" 'nom de l’objet appelant
                    End If
                End If
            End If
NextShape:
        Next shp

        'Boutons de formulaire (si présents)
        For Each obj In ws.Buttons
            If obj.OnAction <> vbNullString Then
                nomMacro = obj.OnAction
                If dictIndex.Exists(nomMacro) Then
                idx = dictIndex(nomMacro)
                    tableProc(idx, 7) = tableProc(idx, 7) + 1
                    If tableProc(idx, 8) = vbNullString Then
                        tableProc(idx, 8) = obj.Name & " (" & ws.Name & ")"
                    Else
                        tableProc(idx, 8) = tableProc(idx, 8) & vbCrLf & obj.Name & " (" & ws.Name & ")"
                    End If
                End If
            End If
        Next obj
    Next ws
    
End Sub

Sub IncrementerAppelsFonctionsCodeDirect(dictIndex As Object, tableProc() As Variant, indexMax As Long)

    Debug.Print "   4. Comptage des appels aux fonctions (dans le code)"
    
    Dim comp As Object, lignes() As String, ligne As String
    Dim i As Long, nomFonction As Variant, valeur As String
    Dim pos1 As Long, pos2 As Long, idx As Long

    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Or comp.Type = vbext_ct_ClassModule Or _
           comp.Type = vbext_ct_MSForm Or comp.Type = vbext_ct_Document Then

            lignes = Split(comp.codeModule.Lines(1, comp.codeModule.CountOfLines), vbCrLf)

            For i = 0 To UBound(lignes)
                ligne = Trim(lignes(i))
                If ligne = vbNullString Or Left(ligne, 1) = "'" Or _
                   LCase(Left(ligne, 4)) = "sub " Or _
                   LCase(Left(ligne, 9)) = "function " Then
                    GoTo LigneSuivante
                End If

'                If ligne = ".Range(""O6"").Value = Fn_ProchainNumeroFacture" Then Stop
                
                For Each nomFonction In dictIndex.keys
                
'                    If nomFonction = "Fn_ProchainNumeroFacture" Then Stop
                    
                    idx = dictIndex(nomFonction)
                    If LCase(tableProc(idx, 2)) <> "function" Then GoTo NextNom
                    
                    ' Appel direct
                    If Fn_AppelDirectFonction(ligne, CStr(nomFonction)) Then
                        tableProc(idx, 5) = tableProc(idx, 5) + 1
                    End If

                    ' Appel préfixé
                    If Fn_AppelDirectFonctionAvecNomModule(ligne, CStr(nomFonction)) Then
                        tableProc(idx, 6) = tableProc(idx, 6) + 1
                    End If

                    ' Appel indirect
                    If (InStr(LCase(ligne), "application.run") > 0 Or _
                        InStr(LCase(ligne), "evaluate(") > 0 Or _
                        InStr(LCase(ligne), "executeexcel4macro") > 0) Then

                        pos1 = InStr(ligne, """")
                        If pos1 > 0 Then
                            pos2 = InStr(pos1 + 1, ligne, """")
                            If pos2 > pos1 Then
                                valeur = Mid(ligne, pos1 + 1, pos2 - pos1 - 1)
                                valeur = Replace(valeur, "()", vbNullString)
                                valeur = Trim(Split(valeur, "!")(UBound(Split(valeur, "!"))))
                                If LCase(valeur) = LCase(nomFonction) Then
                                    tableProc(idx, 7) = tableProc(idx, 7) + 1
                                End If
                            End If
                        End If
                    End If
NextNom:
                Next nomFonction
LigneSuivante:
            Next i
        End If
    Next comp

End Sub

Function Fn_AppelDirectFonction(ligne As String, nomFonction As String) As Boolean

    Dim l As String: l = " " & LCase(Trim(ligne)) & " "
    Dim f As String: f = LCase(nomFonction)

    Fn_AppelDirectFonction = _
        (InStr(l, f & "(") > 0 Or _
         InStr(l, " " & f & " ") > 0 Or _
         InStr(l, "=" & f & " ") > 0 Or _
         InStr(l, " " & f & "=") > 0 Or _
         InStr(l, ".Value = " & f) > 0 Or _
         InStr(l, "& " & f) > 0 Or _
         InStr(l, "or " & f) > 0 Or InStr(l, "and " & f) > 0 Or _
         InStr(l, "(" & f & ")") > 0 Or _
         InStr(l, "if " & f) > 0 Or InStr(l, "then " & f) > 0)
         
End Function
Function Fn_AppelDirectFonctionAvecNomModule(ligne As String, nomFonction As String) As Boolean

    Dim l As String: l = LCase(ligne)
    Dim f As String: f = LCase(nomFonction)

    Fn_AppelDirectFonctionAvecNomModule = (l Like "*." & f & "*" Or _
                       InStr(l, "." & f & "(") > 0)
                       
End Function

Sub ExporterResultatsFeuille(tableProc() As Variant, indexMax As Long) '2025-07-07 @ 09:27

    Debug.Print "   5. Exportation des résultats vers une feuille"

    Application.EnableEvents = False
    
    'Utilisation de la feuille DocAuditVBA (permanente)
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("DocAuditVBA")
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    ws.Range("A1:J" & lastUsedRow).Clear

    With ws
        'Légende interactive
        .Cells(1, 1).Value = "Double-cliquez sur un nom de procédure (colonne A) pour accéder directement au code VBA"
        .Cells(1, 1).Font.Name = "Aptos Narrow"
        .Cells(1, 1).Font.size = 10
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(0, 102, 204) 'Bleu
        .Cells(1, 1).Interior.Color = RGB(235, 247, 255) 'Bleu pâle
        .Cells(1, 1).HorizontalAlignment = xlCenter
        .Cells(1, 1).VerticalAlignment = xlCenter
        .Cells(1, 1).WrapText = True
        .Cells(1, 10).Value = Format$(Now, "yyyy-mm-dd hh:mm:ss")
        .Cells(1, 10).Font.Name = "Aptos Narrow"
        .Cells(1, 10).Font.size = 10
        .Cells(1, 10).Font.Bold = True
        .Cells(1, 10).Font.Color = vbRed
        .Cells(1, 10).HorizontalAlignment = xlCenter
        .Cells(1, 10).VerticalAlignment = xlCenter
       .Rows(1).RowHeight = 30
        'Entêtes
        .Cells(2, 1).Value = "Nom Procédure / Fonction"
        .Cells(2, 2).Value = "Type proc"
        .Cells(2, 3).Value = "Module"
        .Cells(2, 4).Value = "Type Module"
        .Cells(2, 5).Value = "Appels directs"
        .Cells(2, 6).Value = "Appels préfixés"
        .Cells(2, 7).Value = "Appels indirects"
        .Cells(2, 8).Value = "Total appels"
        .Cells(2, 9).Value = "Objet .OnAction"
        .Cells(2, 10).Value = "Non conformité"

        'Contenu
        Dim i As Long
        For i = 1 To indexMax
            .Cells(i + 2, 1).Value = tableProc(i, 1)
            .Cells(i + 2, 2).Value = tableProc(i, 2)
            .Cells(i + 2, 3).Value = tableProc(i, 3)
            .Cells(i + 2, 4).Value = tableProc(i, 4)
            .Cells(i + 2, 5).Value = tableProc(i, 5) 'directs
            .Cells(i + 2, 6).Value = tableProc(i, 6) 'préfixés
            .Cells(i + 2, 7).Value = tableProc(i, 7) 'indirects
            .Cells(i + 2, 8).FormulaR1C1 = "=RC[-3]+RC[-2]+RC[-1]"
            .Cells(i + 2, 9).Value = tableProc(i, 9)
            If tableProc(i, 9) <> vbNullString Then
                .Cells(i + 2, 10).Interior.Color = RGB(255, 230, 230) 'Rouge pâle
            End If
            .Cells(i + 2, 10).Value = tableProc(i, 9)
        Next i

        'Tri multicritère
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("A3:A" & indexMax + 1), Order:=xlAscending
            .SortFields.Add key:=ws.Range("C3:C" & indexMax + 1), Order:=xlAscending
            .SortFields.Add key:=ws.Range("H3:H" & indexMax + 1), Order:=xlDescending
            .SetRange ws.Range("A2:J" & indexMax + 2)
            .Header = xlYes
            .Apply
        End With
    End With
    
    Call DiagnostiquerConformite(ws, tableProc)
    
    'Mise en forme de la feuille
    With ws
        .Columns("A:J").AutoFit
        
        .Cells.VerticalAlignment = xlTop
    
        'Entêtes
        With .Range("A2:J2")
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(0, 102, 204) 'Bleu vif
            .Font.Color = vbWhite
            .Font.Bold = True
            .Font.Italic = True
            .Font.size = 9
        End With
    
        'Colonnes spécifiques
        .Columns("E").HorizontalAlignment = xlCenter
        .Columns("F").HorizontalAlignment = xlCenter
        .Columns("G").HorizontalAlignment = xlCenter
        .Columns("H").HorizontalAlignment = xlCenter
        .Columns("J").HorizontalAlignment = xlCenter
    
        'Filtre sur toutes les colonnes
        .Range("A2:J2").AutoFilter
    
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
        .Cells(lastRow + 1, 1).Value = "R1 - Usage non autorisé de '_' sauf pour événements (_Click, _Change, etc)"
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
        .HeaderMargin = Application.InchesToPoints(0.16)
        .TopMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.16)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.16)
        .RightMargin = Application.InchesToPoints(0.16)
    End With
    
    Application.EnableEvents = True
    
End Sub

Sub DiagnostiquerConformite(ws As Worksheet, tableProc() As Variant) '2025-07-07 @ 09:27

    Dim i As Long
    Dim nom As String
    Dim totalAppels As Long
    Dim suffixesEvenements As Variant
    suffixesEvenements = Array("_Activate", "_AfterUpdate", "_BeforeClose", "_BeforeDoubleClick", _
        "_BeforeRightClick", "_BeforeUpdate", "_Change", "_Click", "_DblClick", "_Enter", "_Exit", _
        "_GotFocus", "_Initialize", "_ItemCheck", "_ItemClick", "_KeyDown", "_KeyUp", "_MouseDown", _
        "_Open", "_QueryClose", "_SelectionChange", "_SheetActivate", "_SheetDeactivate", "_Terminate")
    Dim suffixe As Variant
    Dim estEvenement As Boolean
    
    Dim diagnostics As String
    Dim procName As String
    
    Dim dernLigneUtilisee As Long
    dernLigneUtilisee = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For i = 3 To dernLigneUtilisee
        nom = Trim(ws.Cells(i, 1).Value)
        
        diagnostics = vbNullString

        'R1 - "_" non autorisé sauf si le nom SE TERMINE par un gestionnaire d’événement ou commence par 'zz_'
        If InStr(nom, "_") > 0 And _
            Not Fn_EstSuffixeEvenement(nom) And _
            Not Fn_EstProcedureAutonome(nom) And _
            Not Fn_EstFonction(nom) Then
            diagnostics = diagnostics & "R1,"
        End If

        'R2 - Aucune accent dans les noms de procédures (choix personnel)
        If Fn_ChaineContientAccents(nom) Then diagnostics = diagnostics & "R2,"

        'R3 - Doit commencer pas par une lettre majuscule
        If Left(nom, 1) <> UCase(Left(nom, 1)) And _
            Not Fn_EstProcedureAutonome(nom) And _
            Not Fn_EstPrefixeSpecial(nom) Then
            diagnostics = diagnostics & "R3,"
        End If
        
        'R4 - Doit commencer par un verbe d’action
        If Not Fn_ValiderCommenceParUnVerbe(nom) And _
            Not Fn_EstProcedureAutonome(nom) And _
            Not Fn_EstPrefixeSpecial(nom) And _
            Not Fn_EstFonction(nom) Then
            diagnostics = diagnostics & "R4,"
        End If
        
        'R5 - Procédure n'est jamais appelé par l'application
        totalAppels = ws.Cells(i, 5).Value + ws.Cells(i, 6).Value + ws.Cells(i, 7).Value
        procName = ws.Cells(i, 1)
        If totalAppels < 1 Then
            estEvenement = False
            For Each suffixe In suffixesEvenements
                If Right(procName, Len(suffixe)) = suffixe Then
                    estEvenement = True
                    Exit For
                End If
            Next suffixe
            If Not estEvenement And Not Fn_EstProcedureAutonome(nom) Then
                diagnostics = diagnostics & "R5,"
            End If
        End If

        'Tous les tests ont été passés
        If Right(diagnostics, 1) = "," Then
            diagnostics = Left(diagnostics, Len(diagnostics) - 1)
        End If
        
        If diagnostics <> vbNullString Then
            ws.Cells(i, 10).Value = diagnostics
        End If
        
    Next i
    
End Sub

Function Fn_ChaineContientAccents(texte As String) As Boolean '2025-07-07 @ 09:27

    Dim i As Long, code As Integer
    Const accents As String = "àâéèêîôùûäëïöüçÀÂÉÈÊÎÔÙÛÄËÏÖÜÇ"
    
    For i = 1 To Len(texte)
        If InStr(accents, Mid(texte, i, 1)) > 0 Then
            Fn_ChaineContientAccents = True
            Exit Function
        End If
    Next i
    
    Fn_ChaineContientAccents = False
    
End Function

Function Fn_EstSuffixeEvenement(nom As String) As Boolean '2025-07-07 @ 09:27

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
        Fn_EstSuffixeEvenement = False
        Exit Function
    End If

    'Vérifie que le suffixe correspond à un événement reconnu
    For Each s In suffixes
        If LCase(Right(nom, Len(s))) = LCase(s) Then
            Fn_EstSuffixeEvenement = True
            Exit Function
        End If
    Next s

    Fn_EstSuffixeEvenement = False
    
End Function

Function Fn_EstProcedureAutonome(nom As String) As Boolean

    If Len(nom) > 3 And Left(nom, 3) = "zz_" Then
        Fn_EstProcedureAutonome = True
    Else
        Fn_EstProcedureAutonome = False
    End If

End Function

Function Fn_EstFonction(nom As String) As Boolean

    If Len(nom) > 3 And Left(nom, 3) = "Fn_" Then
        Fn_EstFonction = True
    Else
        Fn_EstFonction = False
    End If

End Function

Function Fn_EstPrefixeSpecial(nom As String) As Boolean

    Fn_EstPrefixeSpecial = False
    
    If Len(nom) > 3 Then
        Select Case Left(nom, 3)
            Case "chk", "cmb", "img", "lst", "shp", "txt"
                Fn_EstPrefixeSpecial = True
        End Select
        
        Select Case Left(nom, 4)
            Case "ctrl"
                Fn_EstPrefixeSpecial = True
        End Select
    End If

End Function

Function Fn_ValiderCommenceParUnVerbe(nom As String) As Boolean '2025-07-07 @ 09:27

    Dim verbesAction As Variant
    verbesAction = Array("Acceder", "Activer", "Actualiser", "Additionner", "Afficher", "Ajouter", "Ajuster", _
                         "Aller", "Analyser", "Annuler", "Appeler", "Appliquer", "Arreter", "Assembler", "Batir", _
                         "Cacher", "Calculer", "Changer", "Charger", "Cocher", "Confirmer", "Comptabiliser", _
                         "Compter", "Comparer", "Configurer", "Connecter", "Construire", "Convertir", "Copier", "Corriger", "Creer", _
                         "Decocher", "Demarrer", "Deplacer", "Detecter", "Determiner", "Detruire", "Diagnostiquer", _
                         "Ecrire", "Effacer", "Enregistrer", "Envoyer", "Evaluer", "Executer", "Exporter", "Extraire", _
                         "Fermer", "Filtrer", "Finaliser", "Fixer", "Formater", "Fusionner", "Gerer", "Generer", "Identifier", _
                         "Importer", "Imprimer", "Incrementer", "Initialiser", "Inserer", "Lire", "Lister", "Marquer", "Mettre", _
                         "Modifier", "Montrer", "Nettoyer", "Noter", "Obtenir", "Organiser", "Ouvrir", "Planifier", "Positionner", _
                         "Preparer", "Previsualiser", "Quitter", "Rafraichir", "Rechercher", "Redefinir", "Redemmarer", "Redimensionner", _
                         "Reinitialiser", "Relancer", "Reorganiser", "Remplir", "Restaurer", "Retourner", "Revenir", "Saisir", _
                         "Sauvegarder", "Scanner", "Selectionner", "Supprimer", "Synchroniser", "Tester", "Traiter", "Transferer", "Trier", _
                         "UserForm", "Valider", "Verifier", "Vider", "Visualiser", "Workbook", "Worksheet", "btn", "chk", "cmb", _
                         "cmd", "ctrl", "shp", "txt", "DEB", "ENC", "FAC", "EJ", "GL", "REGUL", "TEC", "TEST")
    
    Dim v As Variant
    For Each v In verbesAction
        If LCase(Left(nom, Len(v))) = LCase(v) Then
            Fn_ValiderCommenceParUnVerbe = True
            Exit Function
        End If
    Next v
    Fn_ValiderCommenceParUnVerbe = False
    
End Function

Function Fn_AllerVersCode(nomModule As String, Optional nomProcedure As String = vbNullString) As Boolean '2025-07-07 @ 09:27

    On Error GoTo Erreur

    Dim comp As VBComponent
    Dim cm As codeModule
    Dim cpane As CodePane
    Dim startLine As Long, numLines As Long

    'Recherche du module dans le projet VBA actif
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If Trim(comp.Name) = Trim(nomModule) Then Exit For
    Next

    If comp Is Nothing Then GoTo Erreur

    comp.Activate

    'Si aucune procédure demandée, c'est terminé
    If nomProcedure = vbNullString Then
        Fn_AllerVersCode = True
        Exit Function
    End If

    'Recherche de la procédure dans le module
    Set cm = comp.codeModule
    startLine = cm.ProcStartLine(nomProcedure, vbext_pk_Proc)
    If startLine < 1 Then GoTo Erreur

    numLines = cm.ProcCountLines(nomProcedure, vbext_pk_Proc)

    'Saut vers le bon CodePane
    For Each cpane In Application.VBE.CodePanes
        If cpane.codeModule Is cm Then
            With cpane
                .Window.Visible = True
                .SetSelection startLine, 1, startLine, 1
            End With

            Fn_AllerVersCode = True
            Exit Function
        End If
    Next

Erreur:
    Fn_AllerVersCode = False
    
End Function

Sub zz_BatirListeProcEtFoncDansDictionnaire() '2025-07-22 @ 12:01

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
                InStr(ligne, "Sub ") > 0 Then
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

Sub zz_AuditMiseEnFormeConditionnelle() '2025-08-06 @ 08:52

    Dim area As Range, cf As FormatCondition
    Dim iRow As Long, ruleIndex As Long
    Dim dictRules As Object: Set dictRules = CreateObject("Scripting.Dictionary")
    Dim uniqueKey As String, doublon As String
    Dim formule As String, plageAdresse As String
    Dim countNonVides As Long
    
    'Créer la feuille d’audit
    Dim feuilleNom As String
    feuilleNom = "AuditMFC"
    Call EffacerEtRecreerWorksheet(feuilleNom)
    Dim auditWs As Worksheet
    Set auditWs = ThisWorkbook.Sheets(feuilleNom)

    auditWs.Range("A1:G1").Value = Array("Feuille", "Adresse", "Type", "Formule1", "NbCellNonVides", "Doublon", "À Supprimer ?")
    iRow = 2
    
    ' Parcourir toutes les feuilles
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set area = ws.Cells.SpecialCells(xlCellTypeAllFormatConditions)
        On Error GoTo 0
        
        If Not area Is Nothing Then
            For Each area In area.Areas
                For ruleIndex = 1 To area.FormatConditions.count
                    On Error Resume Next
                    Set cf = area.FormatConditions(ruleIndex)
                    If cf Is Nothing Then GoTo SkipRule
                    
                    ' Lecture sécurisée du type
                    Dim t As Variant: t = cf.Type
                    
                    ' Lecture sécurisée de la formule
                    formule = ""
                    If t = xlExpression Or t = xlCellValue Then
                        On Error Resume Next
                        formule = cf.Formula1
                        If Err.Number <> 0 Then formule = "[Erreur lecture Formula1]"
                        On Error GoTo 0
                    Else
                        formule = "[Type non compatible]"
                    End If
                    
                    ' Lecture sécurisée de la plage et CountA
                    On Error Resume Next
                    plageAdresse = cf.AppliesTo.Address
                    countNonVides = WorksheetFunction.CountA(cf.AppliesTo)
                    On Error GoTo 0
                    
                    ' Détection de doublon
                    uniqueKey = ws.Name & "|" & t & "|" & formule & "|" & plageAdresse
                    If dictRules.Exists(uniqueKey) Then
                        doublon = "Oui"
                    Else
                        dictRules.Add uniqueKey, 1
                        doublon = "Non"
                    End If
                    
                    ' Remplir l’audit
                    With auditWs
                        .Cells(iRow, 1).Value = ws.Name
                        .Cells(iRow, 2).Value = plageAdresse
                        .Cells(iRow, 3).Value = t
                        Debug.Print "Contenu formule = [" & formule & "]"
                        Debug.Print "Longueur = " & Len(formule)
'                        On Error Resume Next
                        If Len(formule) = 0 Then
                            auditWs.Cells(iRow, 4).Value = "[Formule vide]"
                        Else
                            auditWs.Cells(iRow, 4).Value = Fn_TexteSecurise(formule)
                        End If
                        If Err.Number <> 0 Then
                            auditWs.Cells(iRow, 4).Value = "[Erreur lors de l’écriture]"
                            Err.Clear
                        End If
'                        On Error GoTo 0
                        .Cells(iRow, 5).Value = countNonVides
                        .Cells(iRow, 6).Value = doublon
                        If countNonVides = 0 Or formule = "" Or formule = "[Erreur lecture Formula1]" Or doublon = "Oui" Then
                            .Cells(iRow, 7).Value = "??"
                        Else
                            .Cells(iRow, 7).Value = ""
                        End If
                    End With
                    
                    iRow = iRow + 1
SkipRule:
                    On Error GoTo 0
                Next ruleIndex
            Next area
        End If
    Next ws
    
    MsgBox "Audit terminé — " & iRow - 2 & " règles analysées", vbInformation
    
End Sub

Function Fn_TexteSecurise(formule As String) As String

    On Error GoTo Erreur

    Dim formuleBrute As String

    'Nettoyage et encodage des guillemets
    formuleBrute = Replace(formule, """", """""")

    'Ajout de guillemets autour, pour usage VBA
    Fn_TexteSecurise = Chr(34) & formuleBrute & Chr(34)
    Exit Function

Erreur:
    Fn_TexteSecurise = "Erreur lors du traitement"
    
End Function

Sub zz_AuditDataValidationsInCells() '2025-08-06 @ 08:59

    'Prepare the result worksheet (wsOutput)
    Call EffacerEtRecreerWorksheet("AuditDataValidations")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("AuditDataValidations")
    wsOutput.Cells(1, 1).Value = "SortKey"
    wsOutput.Cells(1, 2).Value = "Worksheet"
    wsOutput.Cells(1, 3).Value = "CellAddress"
    wsOutput.Cells(1, 4).Value = "ValidationType"
    wsOutput.Cells(1, 5).Value = "Formula1"
    wsOutput.Cells(1, 6).Value = "Formula2"
    wsOutput.Cells(1, 7).Value = "Operator"
    wsOutput.Cells(1, 8).Value = "TimeStamp"
    
    Call CreerEnteteDeFeuille(wsOutput.Range("A1:H1"), RGB(0, 112, 192))
    
    'Create the Array to store results in memory
    Dim arr() As Variant
    ReDim arr(1 To 5000, 1 To 8)
    
    ' Loop through each worksheet in the workbook
    Dim dvType As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim timeStamp As String
    Dim X As Long: X = 1
    Dim xAnalyzed As Long
    For Each ws In ThisWorkbook.Worksheets
        'Loop through each cell in the worksheet
        For Each cell In ws.usedRange
            'Check if the cell has data validation
            xAnalyzed = xAnalyzed + 1

            On Error Resume Next
            dvType = vbNullString
            dvType = cell.Validation.Type
            On Error GoTo 0
            
            If dvType <> vbNullString And dvType <> "0" Then
                'Write the data validation details to the output sheet
                arr(X, 1) = ws.Name & Chr$(0) & cell.Address 'Sort Key
                arr(X, 2) = ws.Name
                arr(X, 3) = cell.Address
                arr(X, 4) = dvType
                Select Case dvType
                    Case "2"
                        arr(X, 4) = "Min/Max"
                    Case "3"
                        arr(X, 4) = "Liste"
                    Case Else
                        arr(X, 4) = dvType
                End Select
                On Error Resume Next
                arr(X, 5) = "'" & cell.Validation.Formula1
                On Error GoTo 0
                
                On Error Resume Next
                arr(X, 6) = "'" & cell.Validation.Formula2
                On Error GoTo 0
                
                On Error Resume Next
                arr(X, 7) = "'" & cell.Validation.Operator
                On Error GoTo 0
                
                timeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
                arr(X, 8) = timeStamp

                'Increment the output row counter
                X = X + 1
            End If
        Next cell
    Next ws

    If X > 1 Then
    
        X = X - 1
        
        Call RedimensionnerTableau2D(arr, X, UBound(arr, 2))
        
        Call TrierTableau2DBubble(arr)
        
        'Array to Worksheet
        Dim outputRow As Long: outputRow = 2
        wsOutput.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
        wsOutput.Range("A:A").EntireColumn.Hidden = True 'Do not show the sortKey
        wsOutput.Columns(4).HorizontalAlignment = xlCenter
        wsOutput.Columns(7).HorizontalAlignment = xlCenter
        wsOutput.Columns(8).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        Dim lastUsedRow As Long
        lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "B").End(xlUp).Row
        Dim j As Long, oldWorksheet As String
        oldWorksheet = wsOutput.Range("B" & lastUsedRow).Value
        For j = lastUsedRow To 2 Step -1
            If wsOutput.Range("B" & j).Value <> oldWorksheet Then
                wsOutput.Rows(j + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                oldWorksheet = wsOutput.Range("B" & j).Value
            End If
        Next j
        
        'Since we might have inserted new row, let's update the lastUsedRow
        lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "B").End(xlUp).Row
        With wsOutput.Range("B2:H" & lastUsedRow)
            On Error Resume Next
            ActiveSheet.Cells.FormatConditions.Delete
            On Error GoTo 0
        
            .FormatConditions.Add Type:=xlExpression, Formula1:= _
                "=ET($B2<>"""";MOD(LIGNE();2)=1)"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
            End With
            .FormatConditions(1).StopIfTrue = False
        End With
        
        wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit

    End If

    'AutoFit the columns for better readability
    wsOutput.Columns.AutoFit
    
    'Result print setup
    lastUsedRow = lastUsedRow + 2
    wsOutput.Range("B" & lastUsedRow).Value = "*** " & Format$(xAnalyzed, "###,##0") & _
                                    " cellules analysées dans l'application ***"
    Dim header1 As String: header1 = "Cells Data Validations"
    Dim header2 As String: header2 = "All worksheets"
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsOutput, wsOutput.Range("B2:H" & lastUsedRow), _
                           header1, _
                           header2, _
                           "$1:$1", _
                           "L")
    
    MsgBox "Data validation list were created in worksheet: " & wsOutput.Name
    
    'Libérer la mémoire
    Set cell = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
End Sub

Sub zz_VerifierControlesAssociesToutesFeuilles() '2025-08-21 @ 08:45

    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Sheets("Feuil4")
    wsOut.Range("A1").CurrentRegion.offset(1).Clear
    Dim r As Long
    
    Dim ws As Worksheet
    Dim shp As Shape
    Dim btn As Object
    Dim macroNameRaw As String
    Dim macroName As String
    Dim vbComp As Object
    Dim codeModule As Object
    Dim ligne As Long
    Dim found As Boolean
    Dim oleObj As OLEObject
    
    With wsOut
    .Range("A1").Value = "Feuille"
    .Range("B1").Value = "Contrôle"
    .Range("C1").Value = "Macro assignée"
    .Range("D1").Value = "Type"
    .Range("E1").Value = "Statut"
    End With
    r = 1 ' Commencer à la ligne 2 pour les données

    ' Parcourir toutes les feuilles du classeur
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print "#079 - Vérification des contrôles sur la feuille : " & ws.Name
        
        ' Vérification des Shapes (Formulaires ou Boutons assignés)
        For Each shp In ws.Shapes
            On Error Resume Next
            macroNameRaw = shp.OnAction
            On Error GoTo 0
            
            If macroNameRaw <> vbNullString Then
                ' Extraire uniquement le nom de la macro après le "!"
                If InStr(1, macroNameRaw, "!") > 0 Then
                    macroName = Split(macroNameRaw, "!")(1)
                Else
                    macroName = macroNameRaw
                End If
                
                ' Vérifier si la macro existe
                found = Fn_VerifierMacroExiste(macroName)
                
                ' Résultat de la vérification
                r = r + 1
                wsOut.Cells(r, 1).Value = ws.Name
                wsOut.Cells(r, 2).Value = shp.Name
                wsOut.Cells(r, 3).Value = macroName
                wsOut.Cells(r, 4).Value = "shape"
                If found Then
                    wsOut.Cells(r, 5).Value = "Valide"
                Else
                    wsOut.Cells(r, 5).Value = "Manquante"
                End If
            End If
        Next shp
        
        ' Vérification des contrôles ActiveX
        For Each oleObj In ws.OLEObjects
            If TypeOf oleObj.Object Is MSForms.CommandButton Then
                ' Construire le nom de la macro à partir du nom du contrôle
                macroName = oleObj.Name & "_Click"
                
                ' Vérifier si la macro existe
                found = Fn_VerifierMacroExiste(macroName, ws.CodeName)
                
                ' Résultat de la vérification
                r = r + 1
                wsOut.Cells(r, 1).Value = ws.Name
                wsOut.Cells(r, 2).Value = oleObj.Name
                wsOut.Cells(r, 3).Value = macroName
                wsOut.Cells(r, 4).Value = "CommandButton"
                If found Then
                    wsOut.Cells(r, 5).Value = "Valide"
                Else
                    wsOut.Cells(r, 5).Value = "Manquante"
                End If
            End If
        Next oleObj
    Next ws

    With wsOut
        .Range("A1").CurrentRegion.Sort Key1:=.Range("A2"), Order1:=xlAscending, _
                                        Key2:=.Range("B2"), Order2:=xlAscending, _
                                        Header:=xlYes
    End With
    
    Call AppliquerZebrage
    
    wsOut.Activate
    
    MsgBox "Vérification terminée sur toutes les feuilles. Consultez la fenêtre Exécution pour les résultats.", vbInformation
    
End Sub

Function Fn_VerifierMacroExiste(macroName As String, Optional moduleName As String = vbNullString) As Boolean '2025-08-21 @ 08:46

    Dim vbComp As Object
    Dim codeModule As Object
    Dim ligne As Long
    Dim nomModule As String
    Dim nomProc As String

    Fn_VerifierMacroExiste = False

    ' Si macroName contient un ".", on le découpe
    If InStr(macroName, ".") > 0 Then
        nomModule = Split(macroName, ".")(0)
        nomProc = Split(macroName, ".")(1)
    Else
        nomProc = macroName
        nomModule = moduleName ' Peut être vide
    End If

    ' Si un module est spécifié, chercher uniquement dedans
    If nomModule <> vbNullString Then
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(nomModule)
        On Error GoTo 0
        If Not vbComp Is Nothing Then
            Set codeModule = vbComp.codeModule
            For ligne = 1 To codeModule.CountOfLines
                If codeModule.ProcOfLine(ligne, vbext_pk_Proc) = nomProc Then
                    Fn_VerifierMacroExiste = True
                    Exit Function
                End If
            Next ligne
        End If
        Exit Function
    End If

    ' Sinon, chercher dans tous les modules
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Set codeModule = vbComp.codeModule
        For ligne = 1 To codeModule.CountOfLines
            If codeModule.ProcOfLine(ligne, vbext_pk_Proc) = nomProc Then
                Fn_VerifierMacroExiste = True
                Exit Function
            End If
        Next ligne
    Next vbComp
End Function

Sub AppliquerZebrage() '2025-08-21 @ 08:46

    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Sheets("Feuil4")
    
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = wsOut.Cells(wsOut.Rows.count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If i Mod 2 = 1 Then
            wsOut.Range("A" & i & ":E" & i).Interior.Color = RGB(240, 240, 240)
        Else
            wsOut.Range("A" & i & ":E" & i).Interior.ColorIndex = xlNone
        End If
    Next i
    
End Sub

