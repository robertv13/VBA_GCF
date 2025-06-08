Attribute VB_Name = "modTEC_Evaluation"
Option Explicit

Public gDictHours As Object 'D�claration globale

Sub TEC_Evaluation_Procedure(cutoffDate As String)

    Dim maxDate As Date
    
    Call TEC_Evaluation_Calcul(cutoffDate, maxDate)
    
    Call TEC_Evaluation_Affichage(cutoffDate, maxDate)
    
End Sub

Sub TEC_Evaluation_Calcul(cutoffDate As String, ByRef maxDate As Date)

    If cutoffDate = "" Then
        Exit Sub
    End If
        
    Call modImport.ImporterTEC
    
    Dim Y As Integer, m As Integer, d As Integer
    Y = year(cutoffDate)
    m = month(cutoffDate)
    d = day(cutoffDate)
    maxDate = DateSerial(Y, m, d)
    
    Dim ws As Worksheet: Set ws = wshTEC_Evaluation
    Dim wsSource As Worksheet: Set wsSource = wsdTEC_Local
    
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
    
    'Transf�re la table en m�moire (arr)
    Dim arr As Variant
    arr = wsSource.Range("A3:P" & lastUsedRow).value
    
    'Grande section
    Dim offset As Long
    
    'Dictionaire pour accumuler les heures par professionnel
    Set gDictHours = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim codeClient As String, profInit As String
    Dim hresNettes As Currency, hresNFact As Currency, hresFact As Currency, hresTEC As Currency
    Dim tableau As Variant
    Dim trancheAge As String
    Dim ageTEC As Long, tecID As Long
    For i = 1 To UBound(arr, 1)
        hresNettes = 0
        hresNFact = 0
        hresFact = 0
        hresTEC = 0
        
        If CDate(arr(i, fTECDate)) <= CDate(maxDate) Then
            tecID = CLng(arr(i, fTECTECID))
            profInit = Format$(arr(i, fTECProfID), "000") & arr(i, fTECProf)
            
            'Cette charge a-t-elle �t� D�truite ?
            If UCase$(arr(i, fTECEstDetruit)) = "FAUX" Then
                hresNettes = arr(i, fTECHeures)
            Else
                hresNettes = 0
            End If
            
            'D�termine si la charge -OU- le client sont non-facturable ?
            codeClient = CStr(arr(i, fTECClientID))
            If UCase$(arr(i, fTECEstFacturable)) = "FAUX" Or Fn_Is_Client_Facturable(codeClient) = False Then
                hresNFact = hresNettes
            Else
                hresFact = hresNettes
            End If
            
            If hresNettes <> hresNFact + hresFact Then Stop
            
            'Cette charge a-t-elle �t� factur�e -OU- Factur�e apr�s la date limite ?
            If UCase$(arr(i, fTECEstFacturee)) = "FAUX" Or CDate(arr(i, fTECDateFacturee)) > CDate(maxDate) Then
                If hresFact > 0 Then
                    hresTEC = hresFact
                Else
                    hresTEC = 0
                End If
            End If
            
            'Avons-nous un TEC diff�rent de 0 ?
            If hresTEC > 0 Then
                ageTEC = maxDate - arr(i, fTECDate)
                'D�termine la trancheAge d'�ge
                Select Case ageTEC
                    Case 0 To 30
                        trancheAge = "- de 30 jours"
                    Case 31 To 60
                        trancheAge = "31 @ 60 jours"
                    Case 61 To 90
                        trancheAge = "61 @ 90 jours"
                    Case Is > 90
                        trancheAge = "+ de 90 jours"
                    Case Else
                        trancheAge = "Non d�fini"
                End Select
                
                If Not gDictHours.Exists(profInit) Then
                    gDictHours.Add profInit, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0), _
                                                  CCur(0), CCur(0), CCur(0), CCur(0), CCur(0), _
                                                  CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
                End If
                tableau = gDictHours(profInit) 'Obtenir le tableau a partir du dictionary
                
                'D�termine la section en fonction du client (GC & VG sont toujours dans la premi�re section)
                If codeClient < "2000" Or arr(i, fTECProfID) = 1 Or arr(i, fTECProfID) = 2 Then
                    offset = 0
                Else
                    offset = 5
                End If
                    
                'Heures pour ce TEC
                tableau(offset + 0) = tableau(offset + 0) + hresTEC
                tableau(10 + 0) = tableau(10 + 0) + hresTEC
                
                'Accumule heures selon l'�ge du TEC
                Select Case trancheAge
                    Case "- de 30 jours"
                        tableau(offset + 1) = tableau(offset + 1) + hresTEC
                        tableau(11) = tableau(11) + hresTEC
                    Case "31 @ 60 jours"
                        tableau(offset + 2) = tableau(offset + 2) + hresTEC
                        tableau(12) = tableau(12) + hresTEC
                  Case "61 @ 90 jours"
                        tableau(offset + 3) = tableau(offset + 3) + hresTEC
                        tableau(13) = tableau(13) + hresTEC
                    Case "+ de 90 jours"
                        tableau(offset + 4) = tableau(offset + 4) + hresTEC
                        tableau(14) = tableau(14) + hresTEC
                    End Select
                gDictHours(profInit) = tableau 'Replacer le tableau dans le dictionnaire
            End If
        End If
    Next i
    
End Sub

Sub TEC_Evaluation_Affichage(cutoffDate As String, maxDate As Date)

    Dim ws  As Worksheet
    Set ws = wshTEC_Evaluation
    
    Dim totalHresTEC As Currency, totalValeurTEC As Currency
    Dim prof As Variant
    Dim strProf As String
    Dim prenom As String, nom As String
    Dim profID As Long
    Dim total(0 To 4) As Currency
    Dim tauxHoraire As Currency
    Dim currentRow As Integer
    currentRow = 6

    Dim valeurTEC As Currency
    Dim i As Integer
    For i = 0 To 10 Step 5
        If i = 0 Then
            ws.Range("D" & currentRow).value = "EXCLUANT les clients '2000' (mais INCLUANT les heures de GC & VG de tous les clients)"
        ElseIf i = 5 Then
            ws.Range("D" & currentRow).value = "SEULEMENT les clients '2000'"
        Else
            ws.Range("D" & currentRow).value = "TOUS LES CLIENTS"
        End If
        ws.Range("D" & currentRow).Font.Bold = True
        Erase total
        totalValeurTEC = 0
        currentRow = currentRow + 1

        Application.EnableEvents = False
        Application.ScreenUpdating = False

        For Each prof In Fn_Sort_Dictionary_By_Keys(gDictHours) 'Sort dictionary by keys in ascending order
            strProf = Mid$(prof, 4)
            profID = Fn_GetID_From_Initials(strProf)
            prenom = Fn_Get_Prenom_From_Initials(strProf)
            nom = Fn_Get_Nom_From_Initials(strProf)
            prenom = prenom & " " & Left$(nom, 1) & "."
            If gDictHours(prof)(i) <> 0 Then
                tauxHoraire = Fn_Get_Hourly_Rate(profID, ws.Range("L3").value)
'                Debug.Print i, prof, gDictHours(prof)(i), tauxHoraire, gDictHours(prof)(i) * tauxHoraire
                ws.Range("E" & currentRow).value = prenom
                ws.Range("F" & currentRow).HorizontalAlignment = xlRight
                ws.Range("F" & currentRow).value = Format$(gDictHours(prof)(i), "#,##0.00")
                ws.Range("G" & currentRow).HorizontalAlignment = xlRight
                ws.Range("G" & currentRow).value = Format$(tauxHoraire, "#,##0.00 $")
                ws.Range("H" & currentRow).HorizontalAlignment = xlRight
                ws.Range("H" & currentRow).value = Format$(gDictHours(prof)(i) * tauxHoraire, "###,##0.00 $")
                ws.Range("I" & currentRow).HorizontalAlignment = xlRight
                ws.Range("I" & currentRow).value = Format$(gDictHours(prof)(i + 1), "#,##0.00")
                ws.Range("J" & currentRow).HorizontalAlignment = xlRight
                ws.Range("J" & currentRow).value = Format$(gDictHours(prof)(i + 2), "#,##0.00")
                ws.Range("K" & currentRow).HorizontalAlignment = xlRight
                ws.Range("K" & currentRow).value = Format$(gDictHours(prof)(i + 3), "#,##0.00")
                ws.Range("L" & currentRow).HorizontalAlignment = xlRight
                ws.Range("L" & currentRow).value = Format$(gDictHours(prof)(i + 4), "#,##0.00")
                currentRow = currentRow + 1
                Dim ii As Integer
                For ii = 0 To 4
                    total(ii) = total(ii) + gDictHours(prof)(i + ii)
                Next ii

                totalValeurTEC = totalValeurTEC + (gDictHours(prof)(i) * tauxHoraire)
                If i = 0 Then
                    valeurTEC = valeurTEC + (gDictHours(prof)(i) * tauxHoraire)
                End If
            End If
        Next prof

        ws.Range("E" & currentRow).HorizontalAlignment = xlLeft
        ws.Range("E" & currentRow & ":L" & currentRow).Font.Bold = True
        ws.Range("F" & currentRow & ":L" & currentRow).HorizontalAlignment = xlRight

        ws.Range("E" & currentRow).value = "* Totaux *"
        ws.Range("F" & currentRow).value = Format$(total(0), "#,##0.00")
        ws.Range("H" & currentRow).value = Format$(totalValeurTEC, "###,##0.00 $")
        If i = 0 Then
            With ws.Range("H" & currentRow).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
             End With
        End If
        ws.Range("I" & currentRow).value = Format$(total(1), "#,##0.00")
        ws.Range("J" & currentRow).value = Format$(total(2), "#,##0.00")
        ws.Range("K" & currentRow).value = Format$(total(3), "#,##0.00")
        ws.Range("L" & currentRow).value = Format$(total(4), "#,##0.00")
        currentRow = currentRow + 2
    Next i

    'Obtenir le solde d'ouverture & les transactions pour le compte TEC au Grand Livre
    Dim solde As Currency
    solde = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("Travaux en cours"), maxDate)

    'Afficher le solde des TEC au Grand Livre
    ws.Range("D3").Font.Bold = True
    ws.Range("D3").Font.size = 12
    ws.Range("D3").Font.Color = vbRed
    Dim message As String
    message = "Le solde au G/L pour les TEC est de " & Format$(solde, "###,##0.00 $")
    If valeurTEC = solde Then
        message = message & ", donc aucune �criture"
    ElseIf valeurTEC > solde Then
        message = message & ", donc un D�bit de " & Format$(valeurTEC - solde, "###,##0.00 $")
    Else
        message = message & ", donc un Cr�dit de " & Format$(valeurTEC - solde, "###,##0.00 $")
    End If
    ws.Range("D3").value = message
    
    If valeurTEC - solde <> 0 Then '2025-06-07 @ 13:58
        ws.Shapes("EcritureGL").Visible = msoTrue
        ws.Range("B2").value = valeurTEC - solde
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ws.Shapes("Impression").Visible = True
    ws.Range("L3").Select

    'Lib�rer la m�moire
    Set gDictHours = Nothing
    Set prof = Nothing
    Set ws = Nothing

End Sub

'CommentOut - 2025-06-07 @ 14:01
'Sub TEC_Evaluation_Procedure_ACO(cutoffDate As String)
'
'    If cutoffDate = "" Then
'        Exit Sub
'    End If
'
'    Call modImport.ImporterTEC
'
'    Dim maxDate As Date
'    Dim Y As Integer, m As Integer, d As Integer
'    Y = year(cutoffDate)
'    m = month(cutoffDate)
'    d = day(cutoffDate)
'    maxDate = DateSerial(Y, m, d)
'
'    Dim ws As Worksheet: Set ws = wshTEC_Evaluation
'    Dim wsSource As Worksheet: Set wsSource = wsdTEC_Local
'
'    Dim lastUsedRow As Long
'    lastUsedRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
'
'    'Transf�re la table en m�moire (arr)
'    Dim arr As Variant
'    arr = wsSource.Range("A3:P" & lastUsedRow).value
'
'    'Grande section
'    Dim offset As Long
'
'    'Dictionaire pour accumuler les heures par professionnel
'    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
'
'    Dim i As Long
'    Dim codeClient As String, profInit As String
'    Dim hresNettes As Currency, hresNFact As Currency, hresFact As Currency, hresTEC As Currency
'    Dim totalHresTEC As Currency, totalValeurTEC As Currency
'    Dim tableau As Variant
'    Dim trancheAge As String
'    Dim ageTEC As Long, tecID As Long
'    For i = 1 To UBound(arr, 1)
'        hresNettes = 0
'        hresNFact = 0
'        hresFact = 0
'        hresTEC = 0
'
'        If CDate(arr(i, fTECDate)) <= CDate(maxDate) Then
'            tecID = CLng(arr(i, fTECTECID))
'            profInit = Format$(arr(i, fTECProfID), "000") & arr(i, fTECProf)
'
'            'Cette charge a-t-elle �t� D�truite ?
'            If UCase$(arr(i, fTECEstDetruit)) = "FAUX" Then
'                hresNettes = arr(i, fTECHeures)
'            Else
'                hresNettes = 0
'            End If
'
'            'D�termine si la charge -OU- le client sont non-facturable ?
'            codeClient = CStr(arr(i, fTECClientID))
'            If UCase$(arr(i, fTECEstFacturable)) = "FAUX" Or Fn_Is_Client_Facturable(codeClient) = False Then
'                hresNFact = hresNettes
'            Else
'                hresFact = hresNettes
'            End If
'
'            If hresNettes <> hresNFact + hresFact Then Stop
'
'            'Cette charge a-t-elle �t� factur�e -OU- Factur�e apr�s la date limite ?
'            If UCase$(arr(i, fTECEstFacturee)) = "FAUX" Or CDate(arr(i, fTECDateFacturee)) > CDate(maxDate) Then
'                If hresFact > 0 Then
'                    hresTEC = hresFact
'                Else
'                    hresTEC = 0
'                End If
'            End If
'
'            'Avons-nous un TEC diff�rent de 0 ?
'            If hresTEC > 0 Then
'                ageTEC = maxDate - arr(i, fTECDate)
'                'D�termine la trancheAge d'�ge
'                Select Case ageTEC
'                    Case 0 To 30
'                        trancheAge = "- de 30 jours"
'                    Case 31 To 60
'                        trancheAge = "31 @ 60 jours"
'                    Case 61 To 90
'                        trancheAge = "61 @ 90 jours"
'                    Case Is > 90
'                        trancheAge = "+ de 90 jours"
'                    Case Else
'                        trancheAge = "Non d�fini"
'                End Select
'
'                If Not dictHours.Exists(profInit) Then
'                    dictHours.Add profInit, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0), _
'                                                  CCur(0), CCur(0), CCur(0), CCur(0), CCur(0), _
'                                                  CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
'                End If
'                tableau = dictHours(profInit) 'Obtenir le tableau a partir du dictionary
'
'                'D�termine la section en fonction du client (GC & VG sont toujours dans la premi�re section)
'                If codeClient < "2000" Or arr(i, fTECProfID) = 1 Or arr(i, fTECProfID) = 2 Then
'                    offset = 0
'                Else
'                    offset = 5
'                End If
'
'                'Heures pour ce TEC
'                tableau(offset + 0) = tableau(offset + 0) + hresTEC
'                tableau(10 + 0) = tableau(10 + 0) + hresTEC
'
'                'Accumule heures selon l'�ge du TEC
'                Select Case trancheAge
'                    Case "- de 30 jours"
'                        tableau(offset + 1) = tableau(offset + 1) + hresTEC
'                        tableau(11) = tableau(11) + hresTEC
'                    Case "31 @ 60 jours"
'                        tableau(offset + 2) = tableau(offset + 2) + hresTEC
'                        tableau(12) = tableau(12) + hresTEC
'                  Case "61 @ 90 jours"
'                        tableau(offset + 3) = tableau(offset + 3) + hresTEC
'                        tableau(13) = tableau(13) + hresTEC
'                    Case "+ de 90 jours"
'                        tableau(offset + 4) = tableau(offset + 4) + hresTEC
'                        tableau(14) = tableau(14) + hresTEC
'                    End Select
'                dictHours(profInit) = tableau 'Replacer le tableau dans le dictionnaire
'            End If
'        End If
'    Next i
'
'    Dim prof As Variant
'    Dim strProf As String
'    Dim prenom As String, nom As String
'    Dim profID As Long
'    Dim total(0 To 4) As Currency
'    Dim tauxHoraire As Currency
'    Dim currentRow As Integer
'    currentRow = 6
'
'    Dim valeurTEC As Currency
'    For i = 0 To 10 Step 5
'        If i = 0 Then
'            ws.Range("D" & currentRow).value = "EXCLUANT les clients '2000' (mais INCLUANT les heures de GC & VG de tous les clients)"
'        ElseIf i = 5 Then
'            ws.Range("D" & currentRow).value = "SEULEMENT les clients '2000'"
'        Else
'            ws.Range("D" & currentRow).value = "TOUS LES CLIENTS"
'        End If
'        ws.Range("D" & currentRow).Font.Bold = True
'        Erase total
'        totalValeurTEC = 0
'        currentRow = currentRow + 1
'
'        Application.EnableEvents = False
'        Application.ScreenUpdating = False
'
'        For Each prof In Fn_Sort_Dictionary_By_Keys(dictHours) 'Sort dictionary by keys in ascending order
'            strProf = Mid$(prof, 4)
'            profID = Fn_GetID_From_Initials(strProf)
'            prenom = Fn_Get_Prenom_From_Initials(strProf)
'            nom = Fn_Get_Nom_From_Initials(strProf)
'            prenom = prenom & " " & Left$(nom, 1) & "."
'            If dictHours(prof)(i + 0) <> 0 Then
'                tauxHoraire = Fn_Get_Hourly_Rate(profID, ws.Range("L3").value)
'                ws.Range("E" & currentRow).value = prenom
'                ws.Range("F" & currentRow).HorizontalAlignment = xlRight
'                ws.Range("F" & currentRow).value = Format$(dictHours(prof)(i + 0), "#,##0.00")
'                ws.Range("G" & currentRow).HorizontalAlignment = xlRight
'                ws.Range("G" & currentRow).value = Format$(tauxHoraire, "#,##0.00 $")
'                ws.Range("H" & currentRow).HorizontalAlignment = xlRight
'                ws.Range("H" & currentRow).value = Format$(dictHours(prof)(i + 0) * tauxHoraire, "###,##0.00 $")
'                ws.Range("I" & currentRow).HorizontalAlignment = xlRight
'                ws.Range("I" & currentRow).value = Format$(dictHours(prof)(i + 1), "#,##0.00")
'                ws.Range("J" & currentRow).HorizontalAlignment = xlRight
'                ws.Range("J" & currentRow).value = Format$(dictHours(prof)(i + 2), "#,##0.00")
'                ws.Range("K" & currentRow).HorizontalAlignment = xlRight
'                ws.Range("K" & currentRow).value = Format$(dictHours(prof)(i + 3), "#,##0.00")
'                ws.Range("L" & currentRow).HorizontalAlignment = xlRight
'                ws.Range("L" & currentRow).value = Format$(dictHours(prof)(i + 4), "#,##0.00")
'                currentRow = currentRow + 1
'                Dim ii As Integer
'                For ii = 0 To 4
'                    total(ii) = total(ii) + dictHours(prof)(i + ii)
'                Next ii
'
'                totalValeurTEC = totalValeurTEC + (dictHours(prof)(i) * tauxHoraire)
'                If i = 0 Then
'                    valeurTEC = valeurTEC + (dictHours(prof)(i) * tauxHoraire)
'                End If
'            End If
'        Next prof
'
'        ws.Range("E" & currentRow).HorizontalAlignment = xlLeft
'        ws.Range("E" & currentRow & ":L" & currentRow).Font.Bold = True
'        ws.Range("F" & currentRow & ":L" & currentRow).HorizontalAlignment = xlRight
'
'        ws.Range("E" & currentRow).value = "* Totaux *"
'        ws.Range("F" & currentRow).value = Format$(total(0), "#,##0.00")
'        ws.Range("H" & currentRow).value = Format$(totalValeurTEC, "###,##0.00 $")
'        If i = 0 Then
'            With ws.Range("H" & currentRow).Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .Color = 65535
'                .TintAndShade = 0
'                .PatternTintAndShade = 0
'             End With
'        End If
'        ws.Range("I" & currentRow).value = Format$(total(1), "#,##0.00")
'        ws.Range("J" & currentRow).value = Format$(total(2), "#,##0.00")
'        ws.Range("K" & currentRow).value = Format$(total(3), "#,##0.00")
'        ws.Range("L" & currentRow).value = Format$(total(4), "#,##0.00")
'        currentRow = currentRow + 2
'    Next i
'
'    'Obtenir le solde d'ouverture & les transactions pour le compte TEC au Grand Livre
'    Dim solde As Double
'    solde = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("Travaux en cours"), maxDate)
'
'    'Afficher le solde des TEC au Grand Livre
'    ws.Range("D3").Font.Bold = True
'    ws.Range("D3").Font.size = 12
'    ws.Range("D3").Font.Color = vbRed
'    Dim message As String
'    message = "Le solde au grand livre pour les TEC est de " & Format$(solde, "###,##0.00 $")
'    If valeurTEC = solde Then
'        message = message & ", donc aucune �criture"
'    ElseIf valeurTEC > solde Then
'        message = message & ", donc un D�bit de " & Format$(valeurTEC - solde, "###,##0.00 $")
'    Else
'        message = message & ", donc un Cr�dit de " & Format$(valeurTEC - solde, "###,##0.00 $")
'    End If
'    ws.Range("D3").value = message
'
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'
'    ws.Shapes("Impression").Visible = True
'    ws.Range("L3").Select
'
'    'Lib�rer la m�moire
'    Set dictHours = Nothing
'    Set prof = Nothing
'    Set ws = Nothing
'    Set wsSource = Nothing
'
'End Sub
'
Sub shp_TEC_Evaluation_Impression_Click()

    Call Evaluation_Apercu_Avant_Impression

End Sub

Sub Evaluation_Apercu_Avant_Impression()

    Dim ws As Worksheet: Set ws = wshTEC_Evaluation
    
    Dim rngToPrint As Range
    Set rngToPrint = ws.Range("C2:M31")
    
    Application.EnableEvents = False

'    'Caract�res pour le rapport
'    With rngToPrint.Font
'        .Name = "Aptos Narrow"
'        .size = 10
'    End With
'
    Application.EnableEvents = True
    
    DoEvents

    Dim header1 As String: header1 = "�valuation des TEC au  " & wshTEC_Evaluation.Range("L3").value
    Dim header2 As String: header2 = ""
    
    Call Simple_Print_Setup(wshTEC_Evaluation, rngToPrint, header1, header2, "$1:$1", "P")

    ws.PrintPreview
    
    'Lib�rer la m�moire
    Set rngToPrint = Nothing
    Set ws = Nothing
    
End Sub

Sub shp_TEC_Evaluation_EcritureGL_Click() '2025-06-08 @ 08:37

    Call TEC_Evaluation_EcritureGL
    
End Sub
    
Sub TEC_Evaluation_EcritureGL() '2025-06-08 @ 08:37
    
    '--- D�clarations ---
    Dim ws As Worksheet
    Dim ajustementTEC As Currency
    Dim glTEC As String, glREVTEC As String
    Dim ecr As cGL_Entry

    '--- Initialisation ---
    Set ws = wshTEC_Evaluation

    ajustementTEC = ws.Range("B2").value
    If ajustementTEC < 0 Then Stop
    glTEC = ObtenirNoGlIndicateur("Travaux en cours")
    glREVTEC = ObtenirNoGlIndicateur("Revenus - Travaux en cours")
    
    '--- Instanciation d'un objet GL_Entry
    Set ecr = New cGL_Entry

    '--- Remplissage des propri�t�s globales
    ecr.DateEcriture = ws.Range("L3").value
    ecr.Description = "Ajustement de la valeur des TEC"
    ecr.Source = ""
    ecr.AutreRemarque = "�criture g�n�r�e par l'application"

    'Ajoute autant de lignes que n�cessaire (d�bit positif, cr�dit n�gatif)
    If ajustementTEC > 0 Then
        ecr.AjouterLigne glTEC, "Travaux en cours", ajustementTEC
        ecr.AjouterLigne glREVTEC, "Revenus - Travaux en cours", -ajustementTEC
    Else
        ecr.AjouterLigne glREVTEC, "Revenus - Travaux en cours", -ajustementTEC
        ecr.AjouterLigne glTEC, "Travaux en cours", ajustementTEC
    End If

    '--- �criture ---
    Call AjouterEcritureGL(ecr)
    
End Sub

Sub shp_TEC_Evaluation_Back_To_TEC_Menu_Click()

    Call TEC_Evaluation_Back_To_TEC_Menu
    
End Sub

Sub TEC_Evaluation_Back_To_TEC_Menu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Evaluation:TEC_Evaluation_Back_To_TEC_Menu", "", 0)
    
    wshTEC_Evaluation.Visible = xlSheetVeryHidden
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select
    
    Call Log_Record("modTEC_Evaluation:TEC_Evaluation_Back_To_TEC_Menu", "", startTime)

End Sub

Sub AjouterEcritureGL(entry As cGL_Entry) '2025-06-08 @ 09:37

    '=== BLOC 1 : �criture dans GCF_BD_MASTER.xslx en utilisant ADO ===
    Dim cn As Object
    Dim rs As Object
    Dim cheminMaster As String
    Dim nextNoEntree As Long
    Dim ts As String
    Dim i As Long
    Dim l As cGL_EntryLine
    Dim sql As String

    On Error GoTo CleanUpADO

    'Chemin du classeur MASTER.xlsx
    cheminMaster = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & "GCF_BD_MASTER.xlsx"
    
    'Ouvre connexion ADO
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cheminMaster & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'D�termine le prochain num�ro d'�criture
    Set rs = cn.Execute("SELECT MAX([NoEntr�e]) AS MaxNo FROM [GL_Trans$]")
    If Not rs.EOF And Not IsNull(rs!MaxNo) Then
        nextNoEntree = rs!MaxNo + 1
    Else
        nextNoEntree = 1
    End If
    entry.NoEcriture = nextNoEntree
    rs.Close
    Set rs = Nothing

    'Timestamp unique pour l'�criture
    ts = Format(Now, "yyyy-mm-dd hh:mm:ss")

    'Ajoute chaque ligne d'�criture dans le classeur MASTER.xlsx
    For i = 1 To entry.Lignes.count
        Set l = entry.Lignes(i)
        sql = "INSERT INTO [GL_Trans$] " & _
              "([NoEntr�e],[Date],[Description],[Source],[NoCompte],[Compte],[D�bit],[Cr�dit],[AutreRemarque],[TimeStamp]) " & _
              "VALUES (" & _
              entry.NoEcriture & "," & _
              "'" & Format(entry.DateEcriture, "yyyy-mm-dd") & "'," & _
              "'" & Replace(entry.Description, "'", "''") & "'," & _
              "'" & Replace(entry.Source, "'", "''") & "'," & _
              "'" & l.NoCompte & "'," & _
              "'" & Replace(l.Description, "'", "''") & "'," & _
              IIf(l.Montant >= 0, Replace(l.Montant, ",", "."), "NULL") & "," & _
              IIf(l.Montant < 0, Replace(-l.Montant, ",", "."), "NULL") & "," & _
              "'" & Replace(entry.AutreRemarque, "'", "''") & "'," & _
              "'" & ts & "'" & _
              ")"
        cn.Execute sql
    Next i

    cn.Close: Set cn = Nothing

    '=== BLOC 2 : �criture dans feuille locale (GL_Trans)
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean, oldCalculation As XlCalculation
    Dim wsLocal As Worksheet, lastRow As Long

    'M�moriser l��tat initial d�Excel
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wsLocal = ThisWorkbook.Sheets("GL_Trans")
    lastRow = wsLocal.Cells(wsLocal.Rows.count, 1).End(xlUp).row

    For i = 1 To entry.Lignes.count
        Set l = entry.Lignes(i)
        With wsLocal
            .Cells(lastRow + i, 1).value = entry.NoEcriture
            .Cells(lastRow + i, 2).value = entry.DateEcriture
            .Cells(lastRow + i, 3).value = entry.Description
            .Cells(lastRow + i, 4).value = entry.Source
            .Cells(lastRow + i, 5).value = l.NoCompte
            .Cells(lastRow + i, 6).value = l.Description
            If l.Montant >= 0 Then
                .Cells(lastRow + i, 7).value = l.Montant
                .Cells(lastRow + i, 8).value = ""
            Else
                .Cells(lastRow + i, 7).value = ""
                .Cells(lastRow + i, 8).value = -l.Montant
            End If
            .Cells(lastRow + i, 9).value = entry.AutreRemarque
            .Cells(lastRow + i, 10).value = ts
        End With
    Next i

    wshTEC_Evaluation.Shapes("EcritureGL").Visible = msoFalse
    MsgBox "L'�criture comptable a �t� compl�t�e avec succ�s", vbInformation, "�criture au Grand Livre"

CleanUpADO:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.state = 1 Then rs.Close
    Set rs = Nothing
    If Not cn Is Nothing Then If cn.state = 1 Then cn.Close
    Set cn = Nothing
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l��criture au G/L : " & Err.Description, vbCritical
    End If
    On Error GoTo 0
    
End Sub
'    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean
'    Dim oldDisplayAlerts As Boolean, oldCalculation As XlCalculation
'    Dim cheminMaster As String
'    Dim wbLocal As Workbook, wsLocal As Worksheet
'    Dim wbMaster As Workbook, wsMaster As Worksheet
'    Dim lastRow As Long, nextNoEntree As Long
'    Dim ts As String
'    Dim i As Long
'    Dim l As cGL_EntryLine
'
'    On Error GoTo CleanUp
'
'    'M�moriser l��tat initial d�Excel
'    oldScreenUpdating = Application.ScreenUpdating
'    oldEnableEvents = Application.EnableEvents
'    oldDisplayAlerts = Application.DisplayAlerts
'    oldCalculation = Application.Calculation
'
'    'Mode silencieux
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Application.DisplayAlerts = False
'    Application.Calculation = xlCalculationManual
'
'    cheminMaster = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & "GCF_BD_MASTER.xlsx"
'    Set wbLocal = ThisWorkbook
'    Set wsLocal = wbLocal.Sheets("GL_Trans")
'
'    'Ouvrir le MASTER
'    Set wbMaster = Workbooks.Open(cheminMaster, ReadOnly:=False)
'    wbMaster.Windows(1).Visible = False
'    Set wsMaster = wbMaster.Sheets("GL_Trans")
'
'    'D�terminer le prochain num�ro d'�criture
'    lastRow = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).row
'    If IsNumeric(wsMaster.Cells(lastRow, 1).value) Then
'        nextNoEntree = wsMaster.Cells(lastRow, 1).value + 1
'    Else
'        nextNoEntree = 1
'    End If
'    entry.NoEcriture = nextNoEntree
'
'    'Timestamp unique pour toutes les lignes de l'�criture
'    ts = Format(Now, "yyyy-mm-dd hh:mm:ss")
'
'    '--- 1. �criture dans MASTER
'    For i = 1 To entry.Lignes.count
'        Set l = entry.Lignes(i)
'        With wsMaster
'            .Cells(lastRow + i, 1).value = entry.NoEcriture
'            .Cells(lastRow + i, 2).value = entry.DateEcriture
'            .Cells(lastRow + i, 3).value = entry.Description
'            .Cells(lastRow + i, 4).value = entry.Source
'            .Cells(lastRow + i, 5).value = l.NoCompte
'            .Cells(lastRow + i, 6).value = l.Description
'            If l.Montant >= 0 Then
'                .Cells(lastRow + i, 7).value = l.Montant
'                .Cells(lastRow + i, 8).value = 0
'            Else
'                .Cells(lastRow + i, 7).value = 0
'                .Cells(lastRow + i, 8).value = -l.Montant
'            End If
'            .Cells(lastRow + i, 9).value = entry.AutreRemarque
'            .Cells(lastRow + i, 10).value = ts
'        End With
'    Next i
'
'    wbMaster.Close SaveChanges:=True
'
'    '--- 2. �criture dans la feuille de l'application (local)
'    lastRow = wsLocal.Cells(wsLocal.Rows.count, 1).End(xlUp).row
'    For i = 1 To entry.Lignes.count
'        Set l = entry.Lignes(i)
'        With wsLocal
'            .Cells(lastRow + i, 1).value = entry.NoEcriture
'            .Cells(lastRow + i, 2).value = entry.DateEcriture
'            .Cells(lastRow + i, 3).value = entry.Description
'            .Cells(lastRow + i, 4).value = entry.Source
'            .Cells(lastRow + i, 5).value = l.NoCompte
'            .Cells(lastRow + i, 6).value = l.Description
'            If l.Montant >= 0 Then
'                .Cells(lastRow + i, 7).value = l.Montant
'                .Cells(lastRow + i, 8).value = 0
'            Else
'                .Cells(lastRow + i, 7).value = 0
'                .Cells(lastRow + i, 8).value = -l.Montant
'            End If
'            .Cells(lastRow + i, 9).value = entry.AutreRemarque
'            .Cells(lastRow + i, 10).value = ts
'        End With
'    Next i
'
'CleanUp:
'    Application.ScreenUpdating = oldScreenUpdating
'    Application.EnableEvents = oldEnableEvents
'    Application.DisplayAlerts = oldDisplayAlerts
'    Application.Calculation = oldCalculation
'    If Err.Number <> 0 Then
'        MsgBox "Erreur lors de l��criture au G/L : " & Err.Description, vbCritical
'    End If
'
'End Sub

