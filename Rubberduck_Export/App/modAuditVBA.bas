Attribute VB_Name = "modAuditVBA"
Option Explicit

Dim RCount(1 To 6) As Long

'Constantes pour les index des types d'appels
Public Const IDX_CALL_SIMPLE As Integer = 1
Public Const IDX_CALL_QUALIFIE As Integer = 2
Public Const IDX_DIRECT_SIMPLE As Integer = 3
Public Const IDX_DIRECT_QUALIFIE As Integer = 4
Public Const IDX_INDIRECT_SIMPLE As Integer = 5
Public Const IDX_INDIRECT_QUALIFIE As Integer = 6
Public Const IDX_EVENT_SIMPLE As Integer = 7
Public Const IDX_EVENT_QUALIFIE As Integer = 8
Public Const NB_TYPES_APPELS As Integer = 8

Sub shpAuditVBAProcedures_Click()

    Call AnalyserTousLesNomsDeProcedures

End Sub

Sub AnalyserTousLesNomsDeProcedures() '2025-08-05 @ 13:52

    Dim dictIndex As Object, dictByProc As Object
    Dim colProc As Collection
    Dim indexMax As Long
    
    Application.ScreenUpdating = False
    
    'Effacer la fenêtre immédiate
    Call EffacerFenetreImmediate
    
    Debug.Print "Audit sur les noms de procédures / fonctions - Début du traitement à " & Format$(Now, "hh:nn:ss")

    Dim i As Long
    For i = LBound(RCount) To UBound(RCount)
        RCount(i) = 0
    Next i
    
    'Étape 1 - Construction du tableau
    Call BatirTableProceduresEtFonctions(colProc, dictIndex, dictByProc, indexMax)

    'Étape 2 - Comptage des appels directs aux procédures dans le code
    Call IncrementerAppelsCodeDirect(dictIndex, dictByProc, colProc)

    'Étape 3 - Comptage des appels indirects aux procédures dans le code
    Call IncrementerAppelsIndirect(dictIndex, dictByProc, colProc)
    
    'Étape 4 : Export vers Excel trié et structuré
    Call ExporterResultatsFeuille(dictIndex, colProc, indexMax)

    Debug.Print "Traitement de " & indexMax & " procédures / fonctions, terminé à " & Format$(Now(), "hh:nn:ss")
    
    Application.ScreenUpdating = True
    
    Call AfficherFeuilleResultats("DocAuditVBA")

End Sub
Sub BatirTableProceduresEtFonctions(colProc As Collection, dictIndex As Object, dictByProc As Object, _
                                                                                         indexMax As Long)
    Debug.Print "   1. Création de la liste des procédures / fonctions"
    Application.StatusBar = "Création de la liste des procédures / fonctions - 1 de 5"
    
    Dim vbComp As Object, vbMod As Object, procInfo As Object
    Dim typeModule As String, ligne As String, nomSubOrFunc As String, typeProc As String
    Dim nbLignes As Long, i As Long, key As String
    
    If dictIndex Is Nothing Then Set dictIndex = CreateObject("Scripting.Dictionary")
    If dictByProc Is Nothing Then Set dictByProc = CreateObject("Scripting.Dictionary")
    If colProc Is Nothing Then Set colProc = New Collection
    
    indexMax = 0
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Set vbMod = vbComp.codeModule
        nbLignes = vbMod.CountOfLines
        
        Select Case vbComp.Type
            Case vbext_ct_StdModule:   typeModule = "3_Module Standard"
            Case vbext_ct_ClassModule: typeModule = "4_Classe"
            Case vbext_ct_MSForm:      typeModule = "2_UserForm"
            Case vbext_ct_Document:    typeModule = "1_Feuille Excel"
            Case Else:                 typeModule = "z_Autre"
        End Select
        
        For i = 1 To nbLignes
            ligne = Trim$(vbMod.Lines(i, 1))
            If Len(ligne) > 0 Then
                typeProc = Fn_DetecterTypeProc(ligne)
                ' === DIAGNOSTIC ===
                If typeProc = "Function" Then
                    Debug.Print "Ligne " & i & " détectée comme Function dans " & vbComp.Name
                    Debug.Print "  Contenu: [" & ligne & "]"
                End If
                ' === FIN DIAGNOSTIC ===
                If Not typeProc = "Autre" Then
                    nomSubOrFunc = vbMod.ProcOfLine(i, vbext_pk_Proc)
                    ' === DIAGNOSTIC 2 ===
                    If typeProc = "Function" Then
                        Debug.Print "  Nom procédure: " & nomSubOrFunc
                        Debug.Print "---"
                    End If
                    ' === FIN DIAGNOSTIC 2 ===
                    If Len(nomSubOrFunc) > 0 Then
                        
                        'Clé composite : Module + Procédure
                        key = LCase$(vbComp.Name & "." & nomSubOrFunc)
                        
                        If Not dictIndex.Exists(key) Then
                            indexMax = indexMax + 1
                            Set procInfo = CreateObject("Scripting.Dictionary")
                            procInfo("NomOriginal") = nomSubOrFunc
                            procInfo("Nom") = LCase$(nomSubOrFunc)
                            procInfo("Module") = vbComp.Name
                            procInfo("TypeMod") = typeModule
                            procInfo("TypeProc") = typeProc
                            procInfo("Call-Simple") = 0
                            procInfo("Call-Qualifié") = 0
                            procInfo("Direct-Simple") = 0
                            procInfo("Direct-Qualifié") = 0
                            procInfo("Indirect-Simple") = 0
                            procInfo("Indirect-Qualifié") = 0
                            procInfo("Event-Simple") = 0
                            procInfo("Event-Qualifié") = 0
                            procInfo("Object") = ""
                            procInfo("NonConformité") = ""
                            
                            colProc.Add procInfo
                            dictIndex.Add key, indexMax
                            
                            'Ajout dans dictByProc
                            Dim procName As String
                            procName = LCase$(nomSubOrFunc)
                            If Not dictByProc.Exists(procName) Then
                                Set dictByProc(procName) = New Collection
                            End If
                            dictByProc(procName).Add vbComp.Name
                        End If
                    End If
                End If
            End If
        Next i
    Next vbComp
    
    Application.StatusBar = False
    
End Sub

Function Fn_DetecterTypeProc(ligne As String) As String
    
    Dim l As String: l = LCase$(Trim$(ligne))
    
    ' === IGNORER LES COMMENTAIRES ===
    If Left$(l, 1) = "'" Then
        Fn_DetecterTypeProc = "Autre"
        Exit Function
    End If
    ' === FIN ===
    
    If l Like "*sub *" Then
        Fn_DetecterTypeProc = "Sub"
    ElseIf l Like "*function *" Then
        Fn_DetecterTypeProc = "Function"
    ElseIf l Like "*property get *" Then
        Fn_DetecterTypeProc = "Property Get"
    ElseIf l Like "*property let *" Then
        Fn_DetecterTypeProc = "Property Let"
    ElseIf l Like "*property set *" Then
        Fn_DetecterTypeProc = "Property Set"
    Else
        Fn_DetecterTypeProc = "Autre"
    End If
    
End Function
Sub IncrementerAppelsCodeDirect(dictIndex As Object, dictByProc As Object, colProc As Collection)

    Debug.Print "   2. Comptage des appels direct via le code"
    Application.StatusBar = "Analyse des appels via le code - 2 de 4"
    
    Dim vbComp As Object
    Dim lignes() As String
    Dim ligne As String, saveLigne As String
    Dim toks As Collection
    Dim i As Long, t As Long
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_ClassModule Or _
           vbComp.Type = vbext_ct_MSForm Or vbComp.Type = vbext_ct_Document Then

            lignes = Split(vbComp.codeModule.Lines(1, vbComp.codeModule.CountOfLines), vbCrLf)

            For i = 0 To UBound(lignes)
                saveLigne = lignes(i)
                ligne = Trim$(lignes(i))
                
                'Ignorer les lignes non pertinentes
                If ligne = vbNullString Or Left$(ligne, 1) = "'" Or _
                   LCase$(Left$(ligne, 12)) = "debug.print " Or _
                   LCase$(Left$(ligne, 7)) = "msgbox " Then
                    GoTo LigneSuivante
                End If

                ' Ignorer les déclarations de Sub/Function
                Dim lowerLigne As String
                lowerLigne = LCase$(ligne)
                If (lowerLigne Like "sub *" Or lowerLigne Like "function *" Or _
                    lowerLigne Like "private sub *" Or lowerLigne Like "public sub *" Or _
                    lowerLigne Like "private function *" Or lowerLigne Like "public function *") And _
                   InStr(ligne, "(") > 0 Then
                    GoTo LigneSuivante
                End If

                ' === TRAITEMENT DES APPELS INDIRECTS D'ABORD ===
                Dim extrait As String
                Dim typeAppel As String
                Dim estQualifie As Boolean
                
                If InStr(lowerLigne, "application.run") > 0 Then
                    extrait = Fn_ExtraireEntreGuillemets(ligne)
                    estQualifie = (InStr(extrait, ".") > 0)
                    typeAppel = IIf(estQualifie, "Indirect-Qualifié", "Indirect-Simple")
                    IncrementerAppel dictIndex, dictByProc, colProc, Fn_NormalizeProcName(extrait), typeAppel
                    GoTo LigneSuivante
                    
                ElseIf InStr(lowerLigne, "application.ontime") > 0 Or InStr(lowerLigne, "application.onkey") > 0 Then
                    extrait = Fn_ExtraireDeuxiemeArgument(ligne)
                    estQualifie = (InStr(extrait, ".") > 0)
                    typeAppel = IIf(estQualifie, "Event-Qualifié", "Event-Simple")
                    IncrementerAppel dictIndex, dictByProc, colProc, Fn_NormalizeProcName(extrait), typeAppel
                    GoTo LigneSuivante
                    
                ElseIf InStr(lowerLigne, "evaluate(") > 0 Or InStr(lowerLigne, "executeexcel4macro") > 0 Then
                    extrait = Fn_ExtraireEntreGuillemets(ligne)
                    estQualifie = (InStr(extrait, ".") > 0)
                    typeAppel = IIf(estQualifie, "Indirect-Qualifié", "Indirect-Simple")
                    IncrementerAppel dictIndex, dictByProc, colProc, Fn_NormalizeProcName(extrait), typeAppel
                    GoTo LigneSuivante
                    
                ElseIf InStr(lowerLigne, ".onaction") > 0 Then
                    extrait = Fn_ExtraireNomOnAction(ligne)
                    estQualifie = (InStr(extrait, ".") > 0)
                    typeAppel = IIf(estQualifie, "Event-Qualifié", "Event-Simple")
                    IncrementerAppel dictIndex, dictByProc, colProc, Fn_NormalizeProcName(extrait), typeAppel
                    GoTo LigneSuivante
                End If

                'TOKENISATION POUR APPELS DIRECTS
                Set toks = Fn_TokenizeSemantic(ligne)

                For t = 1 To toks.count
                    Dim tok As String
                    tok = toks(t)

                    ' Ignorer chaînes et commentaires
                    If Left$(tok, 7) = "STRING:" Or Left$(tok, 8) = "COMMENT:" Then GoTo NextToken
                
                    ' === CAS CALL (avec distinction Simple/Qualifié) ===
                    If tok = "IDENT:call" Then
                        ' Vérifier qu'il y a au moins un token après CALL
                        If t + 1 <= toks.count Then
                            If Left$(toks(t + 1), 6) = "IDENT:" Then
                                ' Vérifier s'il y a assez de tokens pour Call Module.Procedure
                                ' On a besoin de : t+1=Module, t+2=., t+3=Procedure
                                If t + 3 <= toks.count Then
                                    ' Maintenant on peut vérifier en toute sécurité
                                    If toks(t + 2) = "SEP:." Then
                                        If Left$(toks(t + 3), 6) = "IDENT:" Then
                                            ' Call qualifié : Call Module.Procedure
                                            IncrementerAppel dictIndex, dictByProc, colProc, Mid$(toks(t + 3), 7), "Call-Qualifié"
                                            t = t + 3
                                            GoTo NextToken
                                        End If
                                    End If
                                End If
                                
                                ' Si on arrive ici, c'est un Call simple
                                IncrementerAppel dictIndex, dictByProc, colProc, Mid$(toks(t + 1), 7), "Call-Simple"
                                t = t + 1
                            End If
                        End If
                        GoTo NextToken
                    End If
                
                    ' === CAS QUALIFIÉ SANS CALL : Module.Procedure ===
                    If Left$(tok, 6) = "IDENT:" Then
                        If t + 2 <= toks.count Then
                            If toks(t + 1) = "SEP:." Then
                                If Left$(toks(t + 2), 6) = "IDENT:" Then
                                    Dim procQualifie As String
                                    procQualifie = Mid$(toks(t + 2), 7)
                                    
                                    ' Vérifier si c'est un appel
                                    If t + 3 <= toks.count Then
                                        Dim tokSuivant As String
                                        tokSuivant = toks(t + 3)
                                        
                                        If tokSuivant = "SEP:(" Or _
                                           Left$(tokSuivant, 6) = "IDENT:" Or _
                                           Left$(tokSuivant, 7) = "STRING:" Or _
                                           Left$(tokSuivant, 3) = "OP:" Or _
                                           tokSuivant = "SEP:," Then
                                            IncrementerAppel dictIndex, dictByProc, colProc, procQualifie, "Direct-Qualifié"
                                        End If
                                    Else
                                        ' Fin de ligne
                                        IncrementerAppel dictIndex, dictByProc, colProc, procQualifie, "Direct-Qualifié"
                                    End If
                                    GoTo NextToken
                                End If
                            End If
                        End If
                    End If
                
                    ' === CAS SIMPLE SANS CALL : Procedure ===
                    If Left$(tok, 6) = "IDENT:" Then
                        ' Vérifier s'il y a un token suivant
                        If t + 1 <= toks.count Then
                            Dim nextTok As String
                            nextTok = toks(t + 1)
                            
                            ' Ne pas traiter si c'est une qualification (suivi de ".")
                            If nextTok = "SEP:." Then GoTo NextToken
                            
                            ' Appel avec parenthèses : MaProc()
                            If nextTok = "SEP:(" Then
                                IncrementerAppel dictIndex, dictByProc, colProc, Mid$(tok, 7), "Direct-Simple"
                                GoTo NextToken
                            End If
                            
                            ' Appel avec arguments directs : MaProc arg1, arg2
                            If Left$(nextTok, 6) = "IDENT:" Or Left$(nextTok, 7) = "STRING:" Then
                                ' Vérifier que ce n'est pas une assignation
                                If t > 1 Then
                                    If toks(t - 1) <> "OP:=" And toks(t - 1) <> "SEP:." Then
                                        IncrementerAppel dictIndex, dictByProc, colProc, Mid$(tok, 7), "Direct-Simple"
                                    End If
                                Else
                                    IncrementerAppel dictIndex, dictByProc, colProc, Mid$(tok, 7), "Direct-Simple"
                                End If
                            End If
                            
                            ' Appel de fonction sans parenthèses (côté droit de =)
                            If nextTok <> "SEP:(" And _
                               Left$(nextTok, 6) <> "IDENT:" And _
                               Left$(nextTok, 7) <> "STRING:" And _
                               nextTok <> "SEP:," Then
                                
                                If t > 1 Then
                                    If toks(t - 1) = "OP:=" Then
                                        IncrementerAppel dictIndex, dictByProc, colProc, Mid$(tok, 7), "Direct-Simple"
                                    End If
                                End If
                            End If
                            
                        Else
                            ' === NOUVEAU : C'est le dernier token de la ligne ===
                            ' Cas: variable = MaFonction (sans parenthèses, fin de ligne)
                            If t > 1 Then
                                If toks(t - 1) = "OP:=" Then
                                    IncrementerAppel dictIndex, dictByProc, colProc, Mid$(tok, 7), "Direct-Simple"
                                End If
                            End If
                        End If
                    End If
                    
                    ' === NOUVEAU : Appel dans une expression/condition (suivi de comparateur) ===
                    If Left$(tok, 6) = "IDENT:" Then
                        If t + 1 <= toks.count Then
                            Dim tokComp As String
                            tokComp = toks(t + 1)
                            
                            ' Si suivi d'un opérateur de comparaison (=, <>, <, >, <=, >=)
                            If tokComp = "OP:=" Or tokComp = "OP:<>" Or _
                               tokComp = "OP:<" Or tokComp = "OP:>" Or _
                               tokComp = "OP:<=" Or tokComp = "OP:>=" Then
                                
                                ' Vérifier que ce n'est PAS le côté gauche d'une assignation
                                ' (dans ce cas, le token précédent serait un type ou Dim)
                                If t > 1 Then
                                    Dim tokPrec As String
                                    tokPrec = toks(t - 1)
                                    
                                    ' Si précédé de : If, ElseIf, And, Or, Not, (, etc.
                                    ' C'est un appel de fonction dans une condition
                                    If tokPrec = "IDENT:if" Or tokPrec = "IDENT:elseif" Or _
                                       tokPrec = "IDENT:and" Or tokPrec = "IDENT:or" Or _
                                       tokPrec = "IDENT:not" Or tokPrec = "SEP:(" Or _
                                       tokPrec = "SEP:," Then
                                        
                                        ' Ne pas compter si c'est une qualification (Module.)
                                        If t + 2 <= toks.count Then
                                            If toks(t + 1) <> "SEP:." Then
                                                IncrementerAppel dictIndex, dictByProc, colProc, Mid$(tok, 7), "Direct-Simple"
                                            End If
                                        Else
                                            IncrementerAppel dictIndex, dictByProc, colProc, Mid$(tok, 7), "Direct-Simple"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

NextToken:
                Next t

LigneSuivante:
            Next i
        End If
    Next vbComp
    
    Application.StatusBar = False

End Sub

Private Function Fn_ExtraireEntreGuillemets(ligne As String) As String

    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(ligne, """")
    If pos1 > 0 Then
        pos2 = InStr(pos1 + 1, ligne, """")
        If pos2 > pos1 Then
            Fn_ExtraireEntreGuillemets = Mid$(ligne, pos1 + 1, pos2 - pos1 - 1)
        End If
    End If
    
End Function

Private Function Fn_ExtraireNomOnAction(ligne As String) As String

    Dim rhs As String
    
    ' Prendre la partie après le signe =
    rhs = Trim(Mid$(ligne, InStr(ligne, "=") + 1))
    
    ' Retirer les apostrophes éventuelles en début/fin
    If Left(rhs, 1) = "'" Then rhs = Mid(rhs, 2)
    If Right(rhs, 1) = "'" Then rhs = Left(rhs, Len(rhs) - 1)
    
    ' Retirer les guillemets éventuels en début/fin
    If Left(rhs, 1) = """" Then rhs = Mid(rhs, 2)
    If Right(rhs, 1) = """" Then rhs = Left(rhs, Len(rhs) - 1)
    
    ' Couper au premier espace (supprimer les arguments)
    If InStr(rhs, " ") > 0 Then rhs = Left(rhs, InStr(rhs, " ") - 1)
    
    Fn_ExtraireNomOnAction = Fn_NormalizeProcName(rhs)
    
End Function

Private Function Fn_ExtraireDeuxiemeArgument(ligne As String) As String

    ' Extrait le deuxième argument d'un appel Application.OnTime
    ' Exemple : Application.OnTime heurePlanifiee, "MaProc"
    ' Retourne : maproc
    
    Dim pos1 As Long, pos2 As Long, extrait As String
    
    ' Chercher la première virgule (séparateur entre heure et nom de procédure)
    pos1 = InStr(ligne, ",")
    If pos1 = 0 Then Exit Function
    
    ' Chercher le premier guillemet après la virgule
    pos1 = InStr(pos1, ligne, """")
    If pos1 = 0 Then Exit Function
    
    ' Chercher le guillemet de fermeture
    pos2 = InStr(pos1 + 1, ligne, """")
    If pos2 = 0 Then Exit Function
    
    ' Extraire le contenu entre guillemets
    extrait = Mid$(ligne, pos1 + 1, pos2 - pos1 - 1)
    
    ' Nettoyer les apostrophes et guillemets parasites en début/fin
    Do While Left(extrait, 1) = "'" Or Left(extrait, 1) = """"
        extrait = Mid(extrait, 2)
    Loop
    Do While Right(extrait, 1) = "'" Or Right(extrait, 1) = """"
        extrait = Left(extrait, Len(extrait) - 1)
    Loop
    
    ' Normaliser
    Fn_ExtraireDeuxiemeArgument = Fn_NormalizeProcName(extrait)
    
End Function

Private Function Fn_NormalizeProcName(procStr As String) As String
    Dim result As String
    result = Trim$(procStr)
    
    ' Supprimer le nom de classeur entre apostrophes
    If Left$(result, 1) = "'" Then
        Dim posExcl As Long
        posExcl = InStr(result, "'!")
        If posExcl > 0 Then
            result = Mid$(result, posExcl + 2)
        End If
    End If
    
    ' Extraire seulement le dernier identifiant après le dernier "."
    ' Pour "Module.Procedure", retourne "Procedure"
    Dim posDot As Long
    posDot = InStrRev(result, ".")
    If posDot > 0 Then
        result = Mid$(result, posDot + 1)
    End If
    
    ' Nettoyer les caractères spéciaux
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    result = Replace(result, """", "")
    result = Replace(result, "'", "")
    
    Fn_NormalizeProcName = LCase$(Trim$(result))
End Function
Private Sub IncrementerAppel(dictIndex As Object, dictByProc As Object, _
                             colProc As Collection, procName As String, typeAppel As String)
    Dim modules As Collection, key As String, idx As Long, m As Long
    Dim suffixes As Variant, s As Variant
    Dim trouve As Boolean
    
    procName = Fn_NormalizeProcName(procName)
    
    ' Vérifier si la procédure existe
    If Not dictByProc.Exists(procName) Then
        Exit Sub
    End If
    
    ' Récupérer tous les modules où cette procédure existe
    Set modules = dictByProc(procName)
    
    ' Liste des suffixes possibles (pour les Property)
    suffixes = Array("", "_GET", "_LET", "_SET")
    
    ' Incrémenter le compteur pour chaque occurrence de la procédure
    For m = 1 To modules.count
        trouve = False
        
        ' Essayer chaque variante de clé (avec et sans suffixe)
        For Each s In suffixes
            key = LCase$(modules(m) & "." & procName & s)
            
            If dictIndex.Exists(key) Then
                idx = dictIndex(key)
                
                ' Vérifier que le type d'appel existe dans procInfo
                If colProc(idx).Exists(typeAppel) Then
                    colProc(idx)(typeAppel) = colProc(idx)(typeAppel) + 1
                    trouve = True
                    ' NE PAS sortir de la boucle - incrémenter TOUTES les variantes (Get, Let, Set)
                Else
                    Debug.Print "ERREUR: Type d'appel inconnu '" & typeAppel & "' pour '" & key & "'"
                End If
            End If
        Next s
        
        ' Si aucune variante trouvée, afficher l'erreur
        If Not trouve Then
            Debug.Print "Clé composite '" & modules(m) & "." & procName & "' introuvable dans dictIndex (aucune variante)"
        End If
    Next m
    
End Sub
'Function TokenizeSimple(ByVal inputText As String) As Variant
'
'    Dim neutralSeparators As Variant, sep As Variant
'    Dim cleaned As String
'
'    'Séparateurs neutres : on les remplace par des espaces
'    neutralSeparators = Array(",", ";", ":", "+", "-", "/", "\", "|", "?", "@", """")
'
'    cleaned = inputText
'
'    'Remplacement des séparateurs neutres
'    For Each sep In neutralSeparators
'        cleaned = Replace(cleaned, sep, " ")
'    Next sep
'
'    ' NB : on laisse passer les séparateurs sémantiques :
'    '   - Point (.) pour Module.Proc
'    '   - Parenthèses () pour appels de fonctions
'    '   - Apostrophe (') pour commentaires
'    '   - Point d’exclamation (!) pour références Excel
'    ' Ceux-ci seront traités dans la logique de comptage
'
'    ' Split sur l’espace
'    TokenizeSimple = Split(Trim(cleaned))
'
'End Function
'

Public Function Fn_TokenizeSemantic(ByVal inputText As String) As Collection

    Dim toks As New Collection
    Dim buf As String
    Dim ch As String, nextCh As String
    Dim i As Long
    Dim state As Integer
    
    state = 0 ' normal
    
    For i = 1 To Len(inputText)
        ch = Mid$(inputText, i, 1)
        If i < Len(inputText) Then
            nextCh = Mid$(inputText, i + 1, 1)
        Else
            nextCh = vbNullString
        End If
        
        Select Case state
        
            Case 2 ' inComment
                buf = buf & ch
                If i = Len(inputText) Then
                    toks.Add "COMMENT:" & buf
                    buf = ""
                End If
            
            Case 1 ' inString
                buf = buf & ch
                If ch = """" Then
                    If nextCh = """" Then
                        buf = buf & nextCh
                        i = i + 1
                    Else
                        toks.Add "STRING:" & buf
                        buf = ""
                        state = 0
                    End If
                End If
            
            Case 0 ' normal
                Select Case ch
                    Case "'"
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        state = 2
                        buf = ch
                        
                    Case """"
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        state = 1
                        buf = ch
                        
                    Case " ", vbTab
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        
                    Case "("
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        toks.Add "SEP:("
                        
                    Case ")"
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        toks.Add "SEP:)"
                        
                    Case "."
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        toks.Add "SEP:."
                        
                    Case "!"
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        toks.Add "SEP:!"
                        
                    Case ","
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        toks.Add "SEP:,"
                        
                    Case ";"
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        toks.Add "SEP:;"
                        
                    Case ":"
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        If nextCh = "=" Then
                            toks.Add "OP::="
                            i = i + 1
                        Else
                            toks.Add "SEP::"
                        End If
                        
                    Case "+", "-", "*", "/", "\", "^", "&", "="
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        toks.Add "OP:" & ch
                        
                    Case "<"
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        If nextCh = "=" Then
                            toks.Add "OP:<=": i = i + 1
                        ElseIf nextCh = ">" Then
                            toks.Add "OP:<>": i = i + 1
                        Else
                            toks.Add "OP:<"
                        End If
                        
                    Case ">"
                        If Len(buf) > 0 Then toks.Add "IDENT:" & LCase$(buf): buf = ""
                        If nextCh = "=" Then
                            toks.Add "OP:>=": i = i + 1
                        Else
                            toks.Add "OP:>"
                        End If
                        
                    Case Else
                        buf = buf & ch
                End Select
        End Select
    Next i
    
    If Len(buf) > 0 Then
        Select Case state
            Case 2: toks.Add "COMMENT:" & buf
            Case 1: toks.Add "STRING:" & buf
            Case Else: toks.Add "IDENT:" & LCase$(buf)
        End Select
    End If
    
    Set Fn_TokenizeSemantic = toks
End Function

Public Sub zz_TestSemantic()

    Dim toks As Collection
    Dim i As Long
    Set toks = Fn_TokenizeSemantic("modSurveillance.EnregistrerActivite ""CheckBox:"" & chk.Name")
    
    For i = 1 To toks.count
        Debug.Print "tokens(" & (i - 1) & ") = " & toks(i)
    Next i
    
End Sub

Private Function Fn_ExtractProcName(ByVal s As String) As String
    Dim tmp As String
    
    tmp = LCase(Trim(s))
    
    ' Ignorer les lignes vides ou commentaires
    If tmp = vbNullString Or Left(tmp, 1) = "'" Then
        Fn_ExtractProcName = vbNullString
        Exit Function
    End If
    
    ' Retirer le mot-clé CALL
    If Left(tmp, 4) = "call" Then
        tmp = Trim(Mid(tmp, 5))
    End If
    
    ' Retirer les parenthèses éventuelles
    If Right(tmp, 2) = "()" Then
        tmp = Left(tmp, Len(tmp) - 2)
    End If
    
    ' Retirer le préfixe de feuille ou de module
    If InStr(tmp, "!") > 0 Then
        tmp = Mid(tmp, InStr(tmp, "!") + 1)
    ElseIf InStr(tmp, ".") > 0 Then
        tmp = Mid(tmp, InStr(tmp, ".") + 1)
    End If
    
    Fn_ExtractProcName = tmp
End Function

Sub IncrementerAppelsIndirect(dictIndex As Object, dictByProc As Object, colProc As Collection)
    Debug.Print "   3. Comptage des appels aux procédures (via Objets)"
    Application.StatusBar = "Analyse des appels aux procédures/fonctions via des Objets - 3 de 4"
    
    Dim ws As Worksheet, shp As Shape, obj As Object
    Dim cb As CommandBar, ctrl As CommandBarControl
    
    For Each ws In ThisWorkbook.Worksheets
        ' --- Shapes ---
        For Each shp In ws.Shapes
            TraiterOnAction shp.OnAction, dictIndex, dictByProc, colProc, "Shape", shp.Name, ws.Name
        Next shp
        
        ' --- Form Controls (tous ceux qui exposent OnAction) ---
        For Each obj In ws.Buttons
            TraiterOnAction obj.OnAction, dictIndex, dictByProc, colProc, "Button", obj.Name, ws.Name
        Next obj
        For Each obj In ws.DropDowns
            TraiterOnAction obj.OnAction, dictIndex, dictByProc, colProc, "DropDown", obj.Name, ws.Name
        Next obj
        For Each obj In ws.CheckBoxes
            TraiterOnAction obj.OnAction, dictIndex, dictByProc, colProc, "CheckBox", obj.Name, ws.Name
        Next obj
        For Each obj In ws.OptionButtons
            TraiterOnAction obj.OnAction, dictIndex, dictByProc, colProc, "OptionButton", obj.Name, ws.Name
        Next obj
        For Each obj In ws.ListBoxes
            TraiterOnAction obj.OnAction, dictIndex, dictByProc, colProc, "ListBox", obj.Name, ws.Name
        Next obj
        For Each obj In ws.Spinners
            TraiterOnAction obj.OnAction, dictIndex, dictByProc, colProc, "Spinner", obj.Name, ws.Name
        Next obj
        For Each obj In ws.ScrollBars
            TraiterOnAction obj.OnAction, dictIndex, dictByProc, colProc, "ScrollBar", obj.Name, ws.Name
        Next obj
    Next ws
    
    Application.StatusBar = False
    
End Sub
'
'Private Function NettoyerOnAction(ByVal s As String) As String
'
'    s = Trim(LCase(s))
'
'    ' Retirer les apostrophes éventuelles en début/fin
'    If Left(s, 1) = "'" Then s = Mid(s, 2)
'    If Right(s, 1) = "'" Then s = Left(s, Len(s) - 1)
'
'    ' Retirer la partie avant "!" si présente (ex: 'fichier!macro')
'    If InStr(s, "!") > 0 Then
'        s = Split(s, "!")(UBound(Split(s, "!")))
'    End If
'
'    ' Retirer les arguments éventuels (tout ce qui suit un espace)
'    If InStr(s, " ") > 0 Then
'        s = Left(s, InStr(s, " ") - 1)
'    End If
'
'    NettoyerOnAction = Fn_NormalizeProcName(s)
'
'End Function

Private Sub TraiterOnAction(onActionStr As String, dictIndex As Object, dictByProc As Object, _
                            colProc As Collection, typeObjet As String, nomObjet As String, nomFeuille As String)
    
    If onActionStr = vbNullString Then Exit Sub
    
    ' Normaliser et extraire le nom de la procédure
    Dim procName As String
    procName = Fn_NormalizeProcName(onActionStr)
    
    If procName = vbNullString Then Exit Sub
    
    ' Déterminer si c'est qualifié (contient un ".")
    Dim typeAppel As String
    If InStr(onActionStr, ".") > 0 Then
        typeAppel = "Event-Qualifié"
    Else
        typeAppel = "Event-Simple"
    End If
    
    ' Incrémenter via la procédure standard
    IncrementerAppel dictIndex, dictByProc, colProc, procName, typeAppel
    
    ' Enregistrer l'objet qui appelle la procédure
    If Not dictByProc.Exists(procName) Then Exit Sub
    
    Dim modules As Collection, key As String, idx As Long, m As Long
    Set modules = dictByProc(procName)
    
    For m = 1 To modules.count
        key = LCase$(modules(m) & "." & procName)
        If dictIndex.Exists(key) Then
            idx = dictIndex(key)
            
            ' Ajouter l'info de l'objet
            If colProc(idx).Exists("Object") Then
                Dim objInfo As String
                objInfo = typeObjet & ":" & nomObjet & " (" & nomFeuille & ")"
                
                ' === UTILISATION DE vbCrLf pour séparer les objets ===
                If colProc(idx)("Object") = "" Then
                    colProc(idx)("Object") = objInfo
                Else
                    ' Ajouter seulement si pas déjà présent
                    If InStr(colProc(idx)("Object"), objInfo) = 0 Then
                        colProc(idx)("Object") = colProc(idx)("Object") & vbCrLf & objInfo
                    End If
                End If
            End If
        End If
    Next m
    
End Sub
'
'Public Function Fn_AjouterObjet(existing As Variant, _
'                             typeObj As String, _
'                             nomObj As String, _
'                             nomFeuille As String) As String
'    Dim cur As String
'    Dim ajout As String
'
'    cur = CStr(existing)
'    ajout = typeObj & ": " & nomObj & " (" & nomFeuille & ")"
'
'    If Len(cur) = 0 Then
'        Fn_AjouterObjet = ajout
'    Else
'        Fn_AjouterObjet = cur & vbCrLf & ajout
'    End If
'
'End Function
''
''Sub IncrementerAppelsFonctionsCodeDirect(dictIndex As Object, dictByProc As Object, colProc As Collection)
''
''    Debug.Print "   4. Comptage des appels aux fonctions (dans le code)"
''    Application.StatusBar = "Analyse des appels directs dans le code - 4 de 5"
''
''    Dim comp As Object, lignes() As String, ligne As String
''    Dim i As Long, toks As Collection, t As Long
''    Dim procName As String, modules As Collection, key As String, idx As Long, m As Long
''
''    For Each comp In ThisWorkbook.VBProject.VBComponents
''        If comp.Type = vbext_ct_StdModule Or comp.Type = vbext_ct_ClassModule Or _
''           comp.Type = vbext_ct_MSForm Or comp.Type = vbext_ct_Document Then
''
''            lignes = Split(comp.codeModule.Lines(1, comp.codeModule.CountOfLines), vbCrLf)
''
''            For i = 0 To UBound(lignes)
''                ligne = Trim$(lignes(i))
''
''                ' Ignorer les lignes vides, commentaires, déclarations
''                If ligne = vbNullString Or Left$(ligne, 1) = "'" Then GoTo LigneSuivante
''                If LCase$(Left$(ligne, 12)) = "debug.print " Or _
''                   LCase$(Left$(ligne, 7)) = "msgbox " Or _
''                   LCase$(ligne) Like "sub *" Or _
''                   LCase$(ligne) Like "public sub *" Or _
''                   LCase$(ligne) Like "private sub *" Or _
''                   LCase$(ligne) Like "function *" Or _
''                   LCase$(ligne) Like "public function *" Or _
''                   LCase$(ligne) Like "private function *" Then GoTo LigneSuivante
''
''                ' --- Vérification indirecte (Application.Run, Evaluate, ExecuteExcel4Macro) ---
''                If InStr(LCase$(ligne), "application.run") > 0 Or _
''                   InStr(LCase$(ligne), "evaluate(") > 0 Or _
''                   InStr(LCase$(ligne), "executeexcel4macro") > 0 Then
''
''                    Dim extrait As String
''                    extrait = Fn_ExtraireEntreGuillemets(ligne)
''                    procName = Fn_NormalizeProcName(extrait)
''                    If dictByProc.Exists(procName) Then
''                        Set modules = dictByProc(procName)
''                        For m = 1 To modules.count
''                            key = LCase$(modules(m) & "." & procName)
''                            If dictIndex.Exists(key) Then
''                                idx = dictIndex(key)
''                                If LCase$(colProc(idx)("TypeProc")) = "function" Then
''                                    colProc(idx)("Indirect") = colProc(idx)("Indirect") + 1
''                                End If
''                            End If
''                        Next m
''                    End If
''                End If
''
''                ' --- Vérification directe et préfixée par tokenisation ---
''                Set toks = Fn_TokenizeSemantic(ligne)
''                For t = 1 To toks.count
''                    Dim tok As String
''                    tok = toks(t)
''
''                    ' Ignorer chaînes et commentaires
''                    If Left$(tok, 7) = "STRING:" Or Left$(tok, 8) = "COMMENT:" Then GoTo NextToken
''
''                    ' Cas direct: ident suivi de "("
''                    If Left$(tok, 6) = "IDENT:" And t < toks.count Then
''                        If toks(t + 1) = "SEP:(" Then
''                            procName = Fn_NormalizeProcName(Mid$(tok, 7))
''                            If dictByProc.Exists(procName) Then
''                                Set modules = dictByProc(procName)
''                                For m = 1 To modules.count
''                                    key = LCase$(modules(m) & "." & procName)
''                                    If dictIndex.Exists(key) Then
''                                        idx = dictIndex(key)
''                                        If LCase$(colProc(idx)("TypeProc")) = "function" Then
''                                            colProc(idx)("Direct") = colProc(idx)("Direct") + 1
''                                        End If
''                                    End If
''                                Next m
''                            End If
''                        End If
''                    End If
''
''                    ' Cas préfixé: Module.Fonction(
''                    If Left$(tok, 6) = "IDENT:" And t + 3 <= toks.count Then
''                        If toks(t + 1) = "SEP:." And Left$(toks(t + 2), 6) = "IDENT:" And toks(t + 3) = "SEP:(" Then
''                            procName = Fn_NormalizeProcName(Mid$(toks(t + 2), 7))
''                            If dictByProc.Exists(procName) Then
''                                Set modules = dictByProc(procName)
''                                For m = 1 To modules.count
''                                    key = LCase$(modules(m) & "." & procName)
''                                    If dictIndex.Exists(key) Then
''                                        idx = dictIndex(key)
''                                        If LCase$(colProc(idx)("TypeProc")) = "function" Then
''                                            colProc(idx)("Préfixé") = colProc(idx)("Préfixé") + 1
''                                        End If
''                                    End If
''                                Next m
''                            End If
''                        End If
''                    End If
''
''NextToken:
''                Next t
''
''LigneSuivante:
''            Next i
''        End If
''    Next comp
''
''    Application.StatusBar = False
''
''End Sub

'Function Fn_AppelDirectFonction(ligne As String, nomFonction As String) As Boolean
'
'    Dim pos As Long
'    Fn_AppelDirectFonction = False
'
'    pos = InStr(1, LCase(ligne), LCase(nomFonction & "("))
'    If pos > 0 Then
'        Fn_AppelDirectFonction = True
'    End If
'
'End Function
'
Sub ExporterResultatsFeuille(dictIndex As Object, colProc As Collection, indexMax As Long)

    Debug.Print "   4. Exportation des résultats vers une feuille"
    Application.StatusBar = "Exportation des résultats vers la feuille - 4 de 4"
    
    Application.EnableEvents = False
    
    'Utilisation de la feuille DocAuditVBA (permanente)
    Dim ws As Worksheet
    Set ws = Fn_ObtenirOuCreerFeuille("DocAuditVBA")
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    ws.Range("A1:P" & lastUsedRow).Clear  ' Changé de J à P (16 colonnes au lieu de 10)

    'Vérification de conformité avec la mémoire
    Call DiagnostiquerConformite(ws, colProc)
    
    With ws
        'Légende interactive
        'Légende interactive
        .Cells(1, 1).Value = "Double-cliquez sur un nom de procédure (colonne A) pour accéder directement au code VBA"
        .Cells(1, 1).Font.Name = "Aptos Narrow"
        .Cells(1, 1).Font.size = 10
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(0, 102, 204)
        .Cells(1, 1).Interior.Color = RGB(235, 247, 255)
        .Cells(1, 1).HorizontalAlignment = xlCenter
        .Cells(1, 1).VerticalAlignment = xlCenter
        .Cells(1, 1).WrapText = True
        
        .Cells(1, 15).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        .Cells(1, 15).Font.Name = "Aptos Narrow"
        .Cells(1, 15).Font.size = 10
        .Cells(1, 15).Font.Bold = True
        .Cells(1, 15).Font.Color = vbRed
        .Cells(1, 15).HorizontalAlignment = xlCenter
        .Cells(1, 15).VerticalAlignment = xlCenter
        .Rows(1).RowHeight = 30
    
        'Entêtes - INTERVERSION des colonnes 14 et 15
        .Cells(2, 1).Value = "Nom Procédure / Fonction"
        .Cells(2, 2).Value = "Type proc"
        .Cells(2, 3).Value = "Module"
        .Cells(2, 4).Value = "Type Module"
        .Cells(2, 5).Value = "Call Simple"
        .Cells(2, 6).Value = "Call Qualifié"
        .Cells(2, 7).Value = "Direct Simple"
        .Cells(2, 8).Value = "Direct Qualifié"
        .Cells(2, 9).Value = "Indirect Simple"
        .Cells(2, 10).Value = "Indirect Qualifié"
        .Cells(2, 11).Value = "Event Simple"
        .Cells(2, 12).Value = "Event Qualifié"
        .Cells(2, 13).Value = "Total appels"
        .Cells(2, 14).Value = "Non conformité"      ' ? INTERVERTI (était en 15)
        .Cells(2, 15).Value = "Objet .OnAction"     ' ? INTERVERTI (était en 14)

        'Contenu
        Dim arr() As Variant
        Dim i As Long
        'Créer un tableau 2D avec indexMax lignes et 15 colonnes
        ReDim arr(1 To indexMax, 1 To 15)
        
        For i = 1 To indexMax
            arr(i, 1) = colProc(i)("NomOriginal")
            arr(i, 2) = colProc(i)("TypeProc")
            arr(i, 3) = colProc(i)("Module")
            arr(i, 4) = colProc(i)("TypeMod")
            arr(i, 5) = colProc(i)("Call-Simple")
            arr(i, 6) = colProc(i)("Call-Qualifié")
            arr(i, 7) = colProc(i)("Direct-Simple")
            arr(i, 8) = colProc(i)("Direct-Qualifié")
            arr(i, 9) = colProc(i)("Indirect-Simple")
            arr(i, 10) = colProc(i)("Indirect-Qualifié")
            arr(i, 11) = colProc(i)("Event-Simple")
            arr(i, 12) = colProc(i)("Event-Qualifié")
            ' Calcul du total
            arr(i, 13) = colProc(i)("Call-Simple") + colProc(i)("Call-Qualifié") + _
                         colProc(i)("Direct-Simple") + colProc(i)("Direct-Qualifié") + _
                         colProc(i)("Indirect-Simple") + colProc(i)("Indirect-Qualifié") + _
                         colProc(i)("Event-Simple") + colProc(i)("Event-Qualifié")
            arr(i, 14) = colProc(i)("NonConformité")
            arr(i, 15) = colProc(i)("Object")
        Next i

        'Déposer le tableau en une seule opération
        ws.Range("A3").Resize(indexMax, 15).Value = arr

        'Tri multicritère
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("A3:A" & indexMax + 2), Order:=xlAscending
            .SortFields.Add key:=ws.Range("C3:C" & indexMax + 2), Order:=xlAscending
            .SortFields.Add key:=ws.Range("M3:M" & indexMax + 2), Order:=xlDescending  ' Changé de H à M (Total)
            .SetRange ws.Range("A2:O" & indexMax + 2)  ' Changé de J à O
            .Header = xlYes
            .Apply
        End With
    End With
    
    'Mise en forme de la feuille
    With ws
        .Columns("A:O").AutoFit  ' Changé de J à O
        
        .Cells.VerticalAlignment = xlTop
    
        ' === NOUVEAU : Retour à la ligne pour la colonne Object ===
        .Columns("O").WrapText = True
'        .Columns("O").ColumnWidth = 45  ' Largeur fixe raisonnable
        .Columns("O").HorizontalAlignment = xlLeft
        .Columns("O").VerticalAlignment = xlTop  ' Alignement en haut
        .Cells(1, 15).VerticalAlignment = xlCenter
        
        'Entêtes
        With .Range("A2:O2")  ' Changé de J2 à O2
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(0, 102, 204) 'Bleu vif
            .Font.Color = vbWhite
            .Font.Bold = True
            .Font.Italic = True
            .Font.size = 9
        End With
    
        'Colonnes des compteurs centrées (colonnes E à M)
        .Columns("E:N").HorizontalAlignment = xlCenter
        .Columns("O").HorizontalAlignment = xlLeft
    
        'Filtre sur toutes les colonnes
        .Range("A2:O2").AutoFilter
    
        'Volet figé entre ligne 2 et 3
        .Activate
        .Range("B3").Select
        ActiveWindow.FreezePanes = True
    
        'Lignes zébrées bleu pâle/blanc
        For i = 3 To indexMax + 2
            If i Mod 2 = 0 Then
                .Range(.Cells(i, 1), .Cells(i, 15)).Interior.Color = RGB(220, 230, 241)
            Else
                .Range(.Cells(i, 1), .Cells(i, 15)).Interior.ColorIndex = xlNone
            End If
        Next i
        
        .Rows("3:" & indexMax + 2).AutoFit
    
        'Légende des non-conformités à la fin
        Dim lastRow As Long: lastRow = indexMax + 4
        .Cells(lastRow, 1).Value = "Légende des non-conformités pour les " & Trim(Format$(indexMax, "# ##0")) & " procédures/fonctions"
        .Cells(lastRow, 2).Value = "Nb. cas"
        
        .Cells(lastRow + 1, 1).Value = "R1 - Usage non autorisé de '_' sauf pour événements " & _
                                                                            "(_Click, _Change, etc)"
        If RCount(1) Then .Cells(lastRow + 1, 2).Value = Trim(Format$(RCount(1), "# ##0"))
        .Cells(lastRow + 2, 1).Value = "R2 - Le nom contient un caractère accentué"
        If RCount(2) Then .Cells(lastRow + 2, 2).Value = Trim(Format$(RCount(2), "# ##0"))
        .Cells(lastRow + 3, 1).Value = "R3 - Le nom ne commence pas par une majuscule"
        If RCount(3) Then .Cells(lastRow + 3, 2).Value = Trim(Format$(RCount(3), "# ##0"))
        .Cells(lastRow + 4, 1).Value = "R4 - Le nom ne commence pas par un verbe d'action reconnu"
        If RCount(4) Then .Cells(lastRow + 4, 2).Value = Trim(Format$(RCount(4), "# ##0"))
        .Cells(lastRow + 5, 1).Value = "R5 - La procédure/fonction n'est appelée de nulle part"
        If RCount(5) Then .Cells(lastRow + 5, 2).Value = Trim(Format$(RCount(5), "# ##0"))
        .Cells(lastRow + 6, 1).Value = "R6 - Le nom de la fonction ne commence pas par 'Fn_'"
        If RCount(6) Then .Cells(lastRow + 6, 2).Value = Trim(Format$(RCount(6), "# ##0"))
        
        .Range(.Cells(lastRow, 1), .Cells(lastRow + 6, 2)).Font.size = 9
        .Range(.Cells(lastRow, 1), .Cells(lastRow + 6, 2)).Font.Italic = True
        
        .Range(.Cells(lastRow, 2), .Cells(lastRow + 6, 2)).Font.Bold = True
        .Range(.Cells(lastRow, 2), .Cells(lastRow + 6, 2)).HorizontalAlignment = xlCenter
        
    End With
    
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    
        .LeftFooter = Format$(Now, "yyyy-mm-dd") & " " & Format$(Now, "hh:nn:ss")
        
        .CenterFooter = ws.Name
    
        .RightFooter = "Page &P de &N"
    
        'Marges serrées pour optimiser l'espace
        .HeaderMargin = Application.InchesToPoints(0.16)
        .TopMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.16)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.16)
        .RightMargin = Application.InchesToPoints(0.16)
    End With
    
    Application.EnableEvents = True
    
    Application.StatusBar = False
    
End Sub

'Function Fn_DetecterTypeProc(ligne As String) As String
'
'    Dim l As String: l = LCase$(Trim$(ligne))
'
'    If l Like "*sub *" Then
'        Fn_DetecterTypeProc = "Sub"
'    ElseIf l Like "*function *" Then
'        Fn_DetecterTypeProc = "Function"
'    ElseIf l Like "*property get *" Then
'        Fn_DetecterTypeProc = "Property Get"
'    ElseIf l Like "*property let *" Then
'        Fn_DetecterTypeProc = "Property Let"
'    ElseIf l Like "*property set *" Then
'        Fn_DetecterTypeProc = "Property Set"
'    Else
'        Fn_DetecterTypeProc = "Autre"
'    End If
'
'End Function
'
'Function Fn_AppelDirectFonctionAvecNomModule(ligne As String, nomFonction As String) As Boolean
'
'    Dim pos As Long
'    Fn_AppelDirectFonctionAvecNomModule = False
'
'    pos = InStr(1, LCase(ligne), "." & LCase(nomFonction) & "(")
'    If pos > 0 Then
'        Fn_AppelDirectFonctionAvecNomModule = True
'    End If
'
'End Function
'
Sub DiagnostiquerConformite(ws As Worksheet, colProc As Collection)
    Dim nom As String
    Dim i As Long, totalAppels As Long
    
    Dim suffixesEvenements As Variant
    suffixesEvenements = Array("_Activate", "_AfterUpdate", "_BeforeClose", "_BeforeDoubleClick", _
                               "_BeforeRightClick", "_BeforeUpdate", "_Change", "_Click", "_Deactivate", _
                               "_DblClick", "_Enter", "_Exit", "_GotFocus", "_Initialize", "_ItemCheck", _
                               "_ItemClick", "_KeyDown", "_KeyUp", "_MouseDown", "MoveDown", "_Open", _
                               "_QueryClose", "_SelectionChange", "_SheetActivate", "_SheetDeactivate", _
                               "_SheetSelectionChange", "_Terminate")
    
    Dim suffixe As Variant
    Dim estEvenement As Boolean
    
    Dim diagnostics As String
    Dim procName As String
    
    For i = 1 To colProc.count
        nom = colProc(i)("NomOriginal")
        
        diagnostics = vbNullString
        
        'R1 - "_" non autorisé sauf s'il s'agit de code dont le nom SE TERMINE par un événement
        '     ou commence par 'zz_' ou 'TU_'
        If InStr(nom, "_") > 0 And _
            Not Fn_EstSuffixeEvenement(nom) And _
            Not Fn_EstProcedureAutonome(nom) And _
            Not Fn_EstProcedureTestUnitaire(nom) And _
            Not Fn_EstFonction(nom) Then
            diagnostics = diagnostics & "R1,"
            RCount(1) = RCount(1) + 1
        End If
        
        'R2 - Aucun accent dans les noms de procédures
        If Fn_ChaineContientAccents(nom) Then
            diagnostics = diagnostics & "R2,"
            RCount(2) = RCount(2) + 1
        End If
        
        ' R3 - Doit commencer par une majuscule
        If Left$(nom, 1) <> UCase$(Left$(nom, 1)) And _
            Not Fn_EstProcedureAutonome(nom) And _
            Not Fn_EstPrefixeReconnu(nom) And _
            InStr(colProc(i)("TypeProc"), "Property") = 0 Then
            diagnostics = diagnostics & "R3,"
            RCount(3) = RCount(3) + 1
        End If
        
        ' R4 - Doit commencer par un verbe d'action
        If Not Fn_ValiderCommenceParUnVerbe(nom) And _
            Not Fn_EstProcedureAutonome(nom) And _
            Not Fn_EstProcedureTestUnitaire(nom) And _
            Not Fn_EstPrefixeReconnu(nom) And _
            Not Fn_EstSuffixeEvenement(nom) And _
            Not Fn_EstFonction(nom) And _
            InStr(colProc(i)("TypeProc"), "Property") = 0 Then
            diagnostics = diagnostics & "R4,"
            RCount(4) = RCount(4) + 1
        End If
        
        ' R5 - Procédure jamais appelée
        ' === MODIFICATION ICI : Calcul avec les 8 types d'appels ===
        totalAppels = colProc(i)("Call-Simple") + colProc(i)("Call-Qualifié") + _
                      colProc(i)("Direct-Simple") + colProc(i)("Direct-Qualifié") + _
                      colProc(i)("Indirect-Simple") + colProc(i)("Indirect-Qualifié") + _
                      colProc(i)("Event-Simple") + colProc(i)("Event-Qualifié")
        
        procName = colProc(i)("NomOriginal")
        If totalAppels < 1 Then
            estEvenement = False
            For Each suffixe In suffixesEvenements
                If Right$(procName, Len(suffixe)) = suffixe Then
                    estEvenement = True
                    Exit For
                End If
            Next suffixe
            If Not estEvenement And Not Fn_EstProcedureAutonome(nom) Then
                diagnostics = diagnostics & "R5,"
                RCount(5) = RCount(5) + 1
            End If
        End If
        
        ' R6 - Conformité du nom pour les fonctions
        If colProc(i)("TypeProc") = "Function" Then
            If Left$(nom, 3) <> "Fn_" Then
                diagnostics = diagnostics & "R6,"
                RCount(6) = RCount(6) + 1
            End If
        End If
        
        ' Nettoyage de la chaîne
        If Right$(diagnostics, 1) = "," Then
            diagnostics = Left$(diagnostics, Len(diagnostics) - 1)
        End If
        
        ' Écriture du diagnostic dans la Collection
        colProc(i)("NonConformité") = diagnostics
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
                     "_BeforeRightClick", "_BeforeUpdate", "_Change", "_Click", "_Deactivate", _
                     "_DblClick", "_Enter", "_Exit", "_Initialize", "_ItemCheck", "_ItemClick", _
                     "_KeyDown", "_KeyUp", "_MoveDown", "_Open", "_QueryClose", "_SelectionChange", _
                     "_SheetActivate", "_SheetChange", "_SheetDeactivate", "_SheetFollowHyperlink", _
                     "_SheetSelectionChange", "_Terminate")
                     
    Dim s As Variant
    
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

    If Len(nom) > 3 And LCase(Left(nom, 3)) = "zz_" Then
        Fn_EstProcedureAutonome = True
    Else
        Fn_EstProcedureAutonome = False
    End If
    
End Function

Function Fn_EstProcedureTestUnitaire(nom As String) As Boolean

    If Len(nom) > 3 And Left(nom, 3) = "TU_" Then
        Fn_EstProcedureTestUnitaire = True
    Else
        Fn_EstProcedureTestUnitaire = False
    End If
    
End Function

Function Fn_EstFonction(nom As String) As Boolean

    If Len(nom) > 3 And Left(nom, 3) = "Fn_" Then
        Fn_EstFonction = True
    Else
        Fn_EstFonction = False
    End If

End Function

Function Fn_EstPrefixeReconnu(nom As String) As Boolean

    Fn_EstPrefixeReconnu = False
    
    Dim prefixes As Variant
    prefixes = Array("btn", "chk", "cmb", "ctrl", "img", "lst", "mylistbox", "mytextbox", "opt", "shp", "txt", _
                     "workbook")
    
    Dim s As Variant
    
    'Vérifie que le préfixe (s'il existe) correspond à un événement reconnu
    For Each s In prefixes
        If LCase(Left(nom, Len(s))) = LCase(s) Then
            Fn_EstPrefixeReconnu = True
            Exit Function
        End If
    Next s

End Function

Function Fn_ValiderCommenceParUnVerbe(nom As String) As Boolean '2025-07-07 @ 09:27

    Dim verbesAction() As String
    verbesAction = Split("Acceder,Activer,Actualiser,Additionner,Afficher,Ajouter,Ajuster," & _
                         "Aller,Analyser,Annuler,Appeler,Appliquer,Arreter,Assembler,Batir," & _
                         "Cacher,Calculer,Changer,Charger,Cocher,Confirmer,Comptabiliser," & _
                         "Compter,Comparer,Configurer,Connecter,Construire,Convertir,Copier,Corriger,Creer," & _
                         "Decocher,Demarrer,Deplacer,Detecter,Determiner,Detruire,Diagnostiquer,Dormir," & _
                         "Ecrire,Effacer,Enregistrer,Envoyer,Evaluer,Executer,Exporter,Extraire," & _
                         "Fermer,Filtrer,Finaliser,Fixer,Formater,Fusionner,Gerer,Generer,Identifier," & _
                         "Importer,Imprimer,Incrementer,Initialiser,Inserer,Lancer,Lire,Lister,Marquer,Mettre," & _
                         "Modifier,Montrer,Nettoyer,Noter,Obtenir,Organiser,Ouvrir,Planifier,Positionner," & _
                         "Preparer,Previsualiser,Proposer,Proteger,Quitter,Rafraichir,Rechercher,Redefinir,Redemmarer,Redimensionner," & _
                         "Reinitialiser,Relancer,Reorganiser,Remplir,Reporter,Restaurer,Retourner,Revenir,Saisir," & _
                         "Sauvegarder,Scanner,Selectionner,Simuler,Supprimer,Surveiller,Synchroniser,Tester,Traiter,Transferer,Trier," & _
                         "UserForm,Valider,Verifier,Verrouiller,Vider,Visualiser,Workbook,Worksheet,btn,chk,cmb," & _
                         "cmd,ctrl,myListBox,myTextBox,opt,shp,txt,DEB,ENC,FAC,EJ,GL,REGUL,TEC,TEST", ",")
    
    Dim i As Long
    Dim v As Variant
    For i = LBound(verbesAction) To UBound(verbesAction)
        If LCase(Left(nom, Len(verbesAction(i)))) = LCase(verbesAction(i)) Then
            Fn_ValiderCommenceParUnVerbe = True
            Exit Function
        End If
    Next i
    Fn_ValiderCommenceParUnVerbe = False
    
End Function

Function Fn_AllerVersCode(nomModule As String, Optional nomProcedure As String = vbNullString) As Boolean
    
    Dim comp As VBComponent
    Dim cm As codeModule
    Dim cpane As CodePane
    Dim startLine As Long, numLines As Long
    Dim found As Boolean
    
    On Error GoTo erreur
    
    ' Recherche du module
    found = False
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If StrComp(Trim(comp.Name), Trim(nomModule), vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next
    
    If Not found Then GoTo erreur
    
    comp.Activate
    
    ' Si aucune procédure demandée, on s'arrête au module
    If nomProcedure = vbNullString Then
        Fn_AllerVersCode = True
        Exit Function
    End If
    
    ' Recherche de la procédure
    Set cm = comp.codeModule
    On Error Resume Next
    startLine = cm.ProcStartLine(nomProcedure, vbext_pk_Proc)
    numLines = cm.ProcCountLines(nomProcedure, vbext_pk_Proc)
    On Error GoTo erreur
    
    If startLine < 1 Or numLines < 1 Then GoTo erreur
    
    ' Saut vers le bon CodePane
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
    
erreur:
    Fn_AllerVersCode = False
End Function

'Sub zz_BatirListeProcEtFoncDansDictionnaire() '2025-07-22 @ 12:01
'
'    Dim comp As Object
'    Dim codeMod As Object
'    Dim ligne As String
'    Dim nomSub As String
'    Dim typeModule As String
'
'    Dim dictProcFunc As Object
'    Set dictProcFunc = CreateObject("Scripting.Dictionary")
'
'    Dim i As Long
'    For Each comp In ThisWorkbook.VBProject.VBComponents
'        Set codeMod = comp.codeModule
'        Select Case comp.Type
'            Case vbext_ct_StdModule: typeModule = "3_Module Standard"
'            Case vbext_ct_ClassModule: typeModule = "4_Classe"
'            Case vbext_ct_MSForm: typeModule = "2_UserForm"
'            Case vbext_ct_Document: typeModule = "1_Feuille Excel"
'            Case vbext_ct_MSForm: typeModule = "2_UserForm"
'            Case Else: typeModule = "z_Autre"
'        End Select
'
'        For i = 1 To codeMod.CountOfLines
'            ligne = Trim(codeMod.Lines(i, 1))
'            'Exclure les commentaires et les déclarations de fonction ou procédures
'            If Left(ligne, 1) = "'" Or _
'                InStr(ligne, "Function ") > 0 Or _
'                InStr(ligne, "Sub ") > 0 Then
'                GoTo NextLigne
'            End If
'
'            nomSub = codeMod.ProcOfLine(i, vbext_pk_Proc)
'
'            If nomSub <> vbNullString Then
'                If Not dictProcFunc.Exists(nomSub) Then
'                    dictProcFunc.Add nomSub, comp.Name
'                End If
'            End If
'
'NextLigne:
'        Next i
'    Next comp
'
'End Sub
'
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

    On Error GoTo erreur

    Dim formuleBrute As String

    'Nettoyage et encodage des guillemets
    formuleBrute = Replace(formule, """", """""")

    'Ajout de guillemets autour, pour usage VBA
    Fn_TexteSecurise = Chr(34) & formuleBrute & Chr(34)
    Exit Function

erreur:
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
                
                timeStamp = Format$(Now(), "dd/mm/yyyy hh:nn:ss")
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

Public Sub zz_AuditerBlocsPlus3Commentaires() '2025-11-18 @ 08:14

    Dim comp As VBComponent
    Dim lignes() As String
    Dim i As Long, debutBloc As Long
    Dim blocActif As Boolean
    Dim nomProc As String
    Dim listeProcedures As Collection, listeBlocs As Collection
    Set listeProcedures = New Collection
    Set listeBlocs = New Collection

    For Each comp In ThisWorkbook.VBProject.VBComponents
        lignes = Split(comp.codeModule.Lines(1, comp.codeModule.CountOfLines), vbCrLf)
        blocActif = False

        For i = 0 To UBound(lignes)
            If Left(Trim(lignes(i)), 1) = "'" Then
                If Not blocActif Then debutBloc = i
                blocActif = True
            Else
                If blocActif And (i - debutBloc >= 3) Then
                    nomProc = Fn_ExtraireNomProcedure(lignes(debutBloc))
                    If nomProc <> "" Then
                        listeProcedures.Add Left(comp.Name & Space(25), 25) & Right(Space(11) & i - (debutBloc + 1) + 1 & " lignes ", 11) & " - " & nomProc & " " & " " & "(Lignes " & (debutBloc + 1) & " à " & i & ")"
                    Else
                        listeBlocs.Add Left(comp.Name & Space(25), 25) & "Les lignes " & (debutBloc + 1) & " à " & i
                    End If
                End If
                blocActif = False
            End If
        Next i

        'Cas en fin de module
        If blocActif And (UBound(lignes) - debutBloc >= 2) Then
            nomProc = Fn_ExtraireNomProcedure(lignes(debutBloc))
            If nomProc <> "" Then
                listeProcedures.Add Left(comp.Name & Space(25), 25) & Right(Space(11) & i - (debutBloc + 1) + 1 & " lignes ", 11) & " - " & nomProc & " ' (Lignes " & (debutBloc + 1) & " à " & (UBound(lignes) + 1) & ")"
            Else
                listeBlocs.Add Left(comp.Name & Space(25), 25) & "Les lignes " & (debutBloc + 1) & " à " & (UBound(lignes) + 1)
            End If
        End If
    Next comp

    'Tri des procédures
    Dim tableau() As String
    ReDim tableau(1 To listeProcedures.count)
    For i = 1 To listeProcedures.count
        tableau(i) = listeProcedures(i)
    Next i
    Call TrierTableau2D(tableau, LBound(tableau), UBound(tableau))

    ' Affichage
    Debug.Print String(46, "-")
    Debug.Print "Blocs de code avec plus de 2 lignes commentées"
    Debug.Print String(46, "-")
    For i = 1 To UBound(tableau)
        Debug.Print tableau(i)
    Next i

    If listeBlocs.count > 0 Then
        Debug.Print vbCrLf & String(46, "-")
        Debug.Print "Blocs commentés (>= 3 lignes)"
        Debug.Print String(46, "-")
        For i = 1 To listeBlocs.count
            Debug.Print listeBlocs(i)
        Next i
    End If
    Debug.Print String(46, "-")
End Sub

Private Function Fn_ExtraireNomProcedure(ligne As String) As String

    Dim txt As String
    If Left(ligne, 1) = "'" Then
        Fn_ExtraireNomProcedure = ligne
        Exit Function
    End If
    txt = ligne
'    txt = LCase(Replace(ligne, "'", ""))
    If txt Like "*sub *" Or txt Like "*function *" Then
'        txt = Replace(txt, "public", "")
'        txt = Replace(txt, "private", "")
'        txt = Replace(txt, "sub", "")
'        txt = Replace(txt, "function", "")
        txt = Trim(Split(txt, "(")(0))
        Fn_ExtraireNomProcedure = txt
    Else
        Fn_ExtraireNomProcedure = ""
    End If
    
End Function

Private Sub TrierTableau2D(arr() As String, ByVal bas As Long, ByVal haut As Long)

    Dim i As Long, j As Long
    Dim pivot As String, temp As String
    i = bas: j = haut
    pivot = arr((bas + haut) \ 2)
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If bas < j Then Call TrierTableau2D(arr, bas, j)
    If i < haut Then Call TrierTableau2D(arr, i, haut)
    
End Sub

'Utilitaire pour identifier les appels à des nom externes (hors classeur)
Sub zz_ListerNomsExternes() '2025-11-17 @ 15:23

    Dim nm As Name
    Dim formule As String
    Dim nomClasseur As String: nomClasseur = ThisWorkbook.Name

    Debug.Print "=== Noms définis avec références externes ==="
    
    For Each nm In ThisWorkbook.Names
        formule = nm.RefersTo
        
        ' Vérifie si la formule contient un nom de classeur différent
        If InStr(1, formule, ".xlsb", vbTextCompare) > 0 Then
            If InStr(1, formule, nomClasseur, vbTextCompare) = 0 Then
                Debug.Print nm.Name & " ? " & formule
            End If
        End If
    Next nm

    Debug.Print "=== Fin de la liste ==="

End Sub

Sub TU_Fn_ExtractProcName()

    Debug.Print Space(3) & Format$(Now(), "hh:nn:ss") & " - Tests unitaires 'Fn_ExtractProcName'"
    
    Debug.Assert Fn_ExtractProcName("Feuil1!MaProc") = "maproc"
    Debug.Assert Fn_ExtractProcName("Call Init()") = "init"
    Debug.Assert Fn_ExtractProcName("'") = vbNullString
    Debug.Assert Fn_ExtractProcName("VisualiserFacturePDF") = "visualiserfacturepdf"
    
    Debug.Print Space(3) & Format$(Now(), "hh:nn:ss") & " - Tous les tests unitaires 'Fn_ExtractProcName' passés."
    
End Sub

Sub TU_FnAllerVersCode()

    Debug.Print Space(3) & Format$(Now(), "hh:nn:ss") & " - Tests unitaires 'Fn_AllerVersCode'"
    
    Debug.Assert Fn_AllerVersCode("modAuditVBA", "AnalyserTousLesNomsDeProcedures") = True
    Debug.Assert Fn_AllerVersCode("ufSaisieHeures", "txtDate_AfterUpdate") = True
    Debug.Assert Fn_AllerVersCode("modTEC_Saisie", "AjouterLigneTEC") = True
    Debug.Assert Fn_AllerVersCode("Module1", "ProcInconnue") = False
    Debug.Assert Fn_AllerVersCode("ModuleInexistant") = False
    
    Debug.Print Space(3) & Format$(Now(), "hh:nn:ss") & " - Tous les tests unitaires 'Fn_AllerVersCode' passés."
    
End Sub

