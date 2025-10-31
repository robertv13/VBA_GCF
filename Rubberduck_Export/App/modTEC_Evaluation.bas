Attribute VB_Name = "modTEC_Evaluation"
Option Explicit

Public gDictHours As Object 'Déclaration globale

Sub EvaluerValeurTEC(cutoffDate As String)

    Dim maxDate As Date
    
    If Not cutoffDate = vbNullString Then
        Call ViderFeuilleEvaluationTEC
        
        Call CalculerValeurTEC(cutoffDate, maxDate)
        
        Call AfficherValeurTEC(cutoffDate, maxDate)
    End If
    
End Sub

Sub ViderFeuilleEvaluationTEC()

    Dim ws As Worksheet
    Set ws = wshTEC_Evaluation
    
    With ws
        .Range("D3").Value = vbNullString 'Message pour écriture de G/L
        .Range("D6:L28").ClearContents
        .Shapes("shpImpression").Visible = msoFalse
        .Shapes("shpEcritureGL").Visible = msoFalse
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub CalculerValeurTEC(cutoffDate As String, ByRef maxDate As Date)

    If cutoffDate = vbNullString Then
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
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).Row
    
    'Transfère la table en mémoire (arr)
    Dim arr As Variant
    arr = wsSource.Range("A3:P" & lastUsedRow).Value
    
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
            
            'Cette charge a-t-elle été Détruite ?
            If UCase$(arr(i, fTECEstDetruit)) = "FAUX" Then
                hresNettes = arr(i, fTECHeures)
            Else
                hresNettes = 0
            End If
            
            'Détermine si la charge -OU- le client sont non-facturable ?
            codeClient = CStr(arr(i, fTECClientID))
            If UCase$(arr(i, fTECEstFacturable)) = "FAUX" Or Fn_Is_Client_Facturable(codeClient) = False Then
                hresNFact = hresNettes
            Else
                hresFact = hresNettes
            End If
            
            If hresNettes <> hresNFact + hresFact Then Stop
            
            'Cette charge a-t-elle été facturée -OU- Facturée après la date limite ?
            If UCase$(arr(i, fTECEstFacturee)) = "FAUX" Or CDate(arr(i, fTECDateFacturee)) > CDate(maxDate) Then
                If hresFact > 0 Then
                    hresTEC = hresFact
                Else
                    hresTEC = 0
                End If
            End If
            
            'Avons-nous un TEC différent de 0 ?
            If hresTEC > 0 Then
                ageTEC = maxDate - arr(i, fTECDate)
                'Détermine la trancheAge d'âge
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
                        trancheAge = "Non défini"
                End Select
                
                If Not gDictHours.Exists(profInit) Then
                    gDictHours.Add profInit, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0), _
                                                  CCur(0), CCur(0), CCur(0), CCur(0), CCur(0), _
                                                  CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
                End If
                tableau = gDictHours(profInit) 'Obtenir le tableau a partir du dictionary
                
                'Détermine la section en fonction du client (GC & VG sont toujours dans la première section)
                If codeClient < "2000" Or arr(i, fTECProfID) = 1 Or arr(i, fTECProfID) = 2 Then
                    offset = 0
                Else
                    offset = 5
                End If
                
                'Ne plus faire de distinction pour les clients de Michel - 2025-06-13 @ 08:33
                offset = 0
                    
                'Heures pour ce TEC
                tableau(offset + 0) = tableau(offset + 0) + hresTEC
                tableau(10 + 0) = tableau(10 + 0) + hresTEC
                
                'Accumule heures selon l'âge du TEC
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

Sub AfficherValeurTEC(cutoffDate As String, maxDate As Date)

    Dim ws As Worksheet
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
    'Ne plus faire de distinction pour les clients de Michel - 2025-06-13 @ 08:33
    For i = 0 To 0 Step 5
'    For i = 0 To 10 Step 5
        If i = 0 Then
            'Ne plus faire de distinction pour les clients de Michel - 2025-06-13 @ 08:33
            ws.Range("D" & currentRow).Value = "TOUS LES CLIENTS"
'            ws.Range("D" & currentRow).Value = "EXCLUANT les clients '2000' (mais INCLUANT les heures de GC & VG de tous les clients)"
        ElseIf i = 5 Then
            ws.Range("D" & currentRow).Value = "SEULEMENT les clients '2000'"
        Else
            ws.Range("D" & currentRow).Value = "TOUS LES CLIENTS"
        End If
        ws.Range("D" & currentRow).Font.Bold = True
        Erase total
        totalValeurTEC = 0
        currentRow = currentRow + 1

        Application.EnableEvents = False
        Application.ScreenUpdating = False

        For Each prof In Fn_TriDictionnaireParCles(gDictHours) 'Sort dictionary by keys in ascending order
            strProf = Mid$(prof, 4)
            profID = Fn_ProfIDAPartirDesInitiales(strProf)
            prenom = Fn_PrenomAPartirDesInitiales(strProf)
            nom = Fn_NomAPartirDesInitiales(strProf)
            prenom = prenom & " " & Left$(nom, 1) & "."
            If gDictHours(prof)(i) <> 0 Then
                tauxHoraire = Fn_Get_Hourly_Rate(profID, ws.Range("L3").Value)
'                Debug.Print i, prof, gDictHours(prof)(i), tauxHoraire, gDictHours(prof)(i) * tauxHoraire
                ws.Range("E" & currentRow).Value = prenom
                ws.Range("F" & currentRow).HorizontalAlignment = xlRight
                ws.Range("F" & currentRow).Value = Format$(gDictHours(prof)(i), "#,##0.00")
                ws.Range("G" & currentRow).HorizontalAlignment = xlRight
                ws.Range("G" & currentRow).Value = Format$(tauxHoraire, "#,##0.00 $")
                ws.Range("H" & currentRow).HorizontalAlignment = xlRight
                ws.Range("H" & currentRow).Value = Format$(gDictHours(prof)(i) * tauxHoraire, "###,##0.00 $")
                ws.Range("I" & currentRow).HorizontalAlignment = xlRight
                ws.Range("I" & currentRow).Value = Format$(gDictHours(prof)(i + 1), "#,##0.00")
                ws.Range("J" & currentRow).HorizontalAlignment = xlRight
                ws.Range("J" & currentRow).Value = Format$(gDictHours(prof)(i + 2), "#,##0.00")
                ws.Range("K" & currentRow).HorizontalAlignment = xlRight
                ws.Range("K" & currentRow).Value = Format$(gDictHours(prof)(i + 3), "#,##0.00")
                ws.Range("L" & currentRow).HorizontalAlignment = xlRight
                ws.Range("L" & currentRow).Value = Format$(gDictHours(prof)(i + 4), "#,##0.00")
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

        ws.Range("E" & currentRow).Value = "* Totaux *"
        ws.Range("F" & currentRow).Value = Format$(total(0), "#,##0.00")
        ws.Range("H" & currentRow).Value = Format$(totalValeurTEC, "###,##0.00 $")
        If i = 0 Then
            With ws.Range("H" & currentRow).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
             End With
        End If
        ws.Range("I" & currentRow).Value = Format$(total(1), "#,##0.00")
        ws.Range("J" & currentRow).Value = Format$(total(2), "#,##0.00")
        ws.Range("K" & currentRow).Value = Format$(total(3), "#,##0.00")
        ws.Range("L" & currentRow).Value = Format$(total(4), "#,##0.00")
        currentRow = currentRow + 2
    Next i

    'Obtenir le solde au G/L pour le compte TEC avec Fn_SoldesParCompteAvecADO - 2025-08-03 @ 09:23
    Dim glTEC As String
    glTEC = Fn_NoCompteAPartirIndicateurCompte("Travaux en cours")
    Dim dictSoldes As Object
    Set dictSoldes = CreateObject("Scripting.Dictionary")
    Set dictSoldes = modGL_Stuff.Fn_SoldesParCompteAvecADO(glTEC, "", maxDate, True)
    Dim solde As Currency
    solde = dictSoldes(glTEC)
    
    'Afficher le solde des TEC au Grand Livre
    ws.Range("D3").Font.Bold = True
    ws.Range("D3").Font.size = 12
    ws.Range("D3").Font.Color = vbRed
    
    Dim message As String
    message = "Le solde au G/L pour les TEC est de " & Format$(solde, "###,##0.00 $")
    If valeurTEC = solde Then
        message = message & ", donc aucune écriture"
    ElseIf valeurTEC > solde Then
        message = message & ", donc un Débit de " & Format$(valeurTEC - solde, "###,##0.00 $")
    Else
        message = message & ", donc un Crédit de " & Format$(valeurTEC - solde, "###,##0.00 $")
    End If
    ws.Range("D3").Value = message
    
    If valeurTEC - solde <> 0 Then '2025-06-07 @ 13:58
        ws.Shapes("shpEcritureGL").Visible = msoTrue
        ws.Range("B2").Value = valeurTEC - solde
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ws.Shapes("shpImpression").Visible = True
    ws.Range("L3").Select

    'Libérer la mémoire
    Set gDictHours = Nothing
    Set prof = Nothing
    Set ws = Nothing

End Sub

Sub shpImprimerValeurTEC_Click()

    Call PrevisualiserRapport

End Sub

Sub PrevisualiserRapport()

    Dim ws As Worksheet: Set ws = wshTEC_Evaluation
    
    Dim rngToPrint As Range
    Set rngToPrint = ws.Range("C2:M31")
    
    Application.EnableEvents = False

    Application.EnableEvents = True
    
    DoEvents

    Dim header1 As String: header1 = "Évaluation des TEC au  " & wshTEC_Evaluation.Range("L3").Value
    Dim header2 As String: header2 = vbNullString
    
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wshTEC_Evaluation, rngToPrint, header1, header2, "$1:$1", "P")

    ws.PrintPreview
    
    'Libérer la mémoire
    Set rngToPrint = Nothing
    Set ws = Nothing
    
End Sub

Sub shpComptabiliserValeurTEC_Click() '2025-06-08 @ 08:37

    Call ComptabiliserValeurTEC
    
End Sub
    
Sub ComptabiliserValeurTEC() '2025-06-08 @ 08:37
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Evaluation:ComptabiliserValeurTEC", vbNullString, 0)
    
    'Déclarations
    Dim ws As Worksheet
    Dim ajustementTEC As Currency
    Dim glTEC As String, glREVTEC As String
    Dim ecr As clsGL_Entry

    'Initialisation
    Set ws = wshTEC_Evaluation

    ajustementTEC = ws.Range("B2").Value
    glTEC = Fn_NoCompteAPartirIndicateurCompte("Travaux en cours")
    glREVTEC = Fn_NoCompteAPartirIndicateurCompte("Revenus - Travaux en cours")
    
    'Instanciation d'un objet GL_Entry
    Set ecr = New clsGL_Entry

    'Remplissage des propriétés globales
    ecr.DateEcriture = ws.Range("L3").Value
    ecr.description = "Ajustement de la valeur des TEC"
    ecr.source = vbNullString

    'Ajoute autant de lignes que nécessaire (débit positif, crédit négatif)
    If ajustementTEC > 0 Then
        ecr.AjouterLigne glTEC, "Travaux en cours", ajustementTEC, "Écriture générée par l'application"
        ecr.AjouterLigne glREVTEC, "Revenus - Travaux en cours", -ajustementTEC, "Écriture générée par l'application"
    Else
        ecr.AjouterLigne glREVTEC, "Revenus - Travaux en cours", -ajustementTEC, "Écriture générée par l'application"
        ecr.AjouterLigne glTEC, "Travaux en cours", ajustementTEC, "Écriture générée par l'application"
    End If

    'Écriture
    Call modGL_Stuff.AjouterEcritureGLADOPlusLocale(ecr, True)
    
    wshTEC_Evaluation.Shapes("shpEcritureGL").Visible = msoFalse
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Evaluation:ComptabiliserValeurTEC", vbNullString, startTime)

End Sub

Sub shpRetournerMenu_Click()

    Call RetournerAuMenu

End Sub

Sub RetournerAuMenu()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Evaluation:RetournerMenu", vbNullString, 0)
    
    Call modAppli.QuitterFeuillePourMenu(wshMenuTEC, True)
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Evaluation:RetournerMenu", vbNullString, startTime)

End Sub


