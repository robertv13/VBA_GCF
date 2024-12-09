Attribute VB_Name = "modTEC_Evaluation"
Option Explicit

Sub TEC_Evaluation_Procedure(cutoffDate As String)

    If cutoffDate = "" Then
        Exit Sub
    End If
        
    Call TEC_Import_All
    
    Dim maxDate As Date
    Dim Y As Integer, m As Integer, d As Integer
    Y = year(cutoffDate)
    m = month(cutoffDate)
    d = day(cutoffDate)
    maxDate = DateSerial(Y, m, d)
    
    Dim ws As Worksheet: Set ws = wshTEC_Evaluation
    Dim wsSource As Worksheet: Set wsSource = wshTEC_Local
    
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
    
    'Transfère la table en mémoire (arr)
    Dim arr As Variant
    arr = wsSource.Range("A3:P" & lastUsedRow).value
    
    'Grande section
    Dim offset As Long
    
    'Dictionaire pour accumuler les heures par professionnel
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim codeClient As String, profInit As String
    Dim hresNettes As Currency, hresNFact As Currency, hresFact As Currency, hresTEC As Currency
    Dim totalHresTEC As Currency, totalValeurTEC As Currency
    Dim tableau As Variant
    Dim trancheAge As String
    Dim ageTEC As Long, TECID As Long
    For i = 1 To UBound(arr, 1)
        hresNettes = 0
        hresNFact = 0
        hresFact = 0
        hresTEC = 0
        
        If CDate(arr(i, 4)) <= CDate(maxDate) Then
            TECID = CLng(arr(i, 1))
'            If TECID = 1755 Or TECID = 1760 Then Stop
            profInit = Format$(arr(i, 2), "000") & arr(i, 3)
            
            'Cette charge a-t-elle été Détruite ?
            If UCase(arr(i, 14)) = "FAUX" Then
                hresNettes = arr(i, 8)
            Else
                hresNettes = 0
            End If
            
            'Détermine si la charge -OU- le client sont non-facturable ?
            codeClient = CStr(arr(i, 5))
            If UCase(arr(i, 10)) = "FAUX" Or Fn_Is_Client_Facturable(codeClient) = False Then
                hresNFact = hresNettes
            Else
                hresFact = hresNettes
            End If
            
            If hresNettes <> hresNFact + hresFact Then Stop
            
            'Cette charge a-t-elle été facturée -OU- Facturée après la date limite ?
            If UCase(arr(i, 12)) = "FAUX" Or Int(arr(i, 13)) > maxDate Then
                If hresFact > 0 Then
                    hresTEC = hresFact
                Else
                    hresTEC = 0
                End If
            End If
            
            'Avons-nous un TEC différent de 0 ?
            If hresTEC > 0 Then
                ageTEC = maxDate - arr(i, 4)
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
                
                If Not dictHours.Exists(profInit) Then
                    dictHours.Add profInit, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0), _
                                                  CCur(0), CCur(0), CCur(0), CCur(0), CCur(0), _
                                                  CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
                End If
                tableau = dictHours(profInit) 'Obtenir le tableau a partir du dictionary
                
                'Détermine la section en fonction du client
                If codeClient < "2000" Then
                    offset = 0
                Else
                    offset = 5
                End If
                    
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
                dictHours(profInit) = tableau 'Replacer le tableau dans le dictionnaire
            
'                Feuil1.Range("A" & TECID).value = TECID
'                Feuil1.Range("B" & TECID).value = hresTEC
                
            End If
        End If
    Next i
    
    Dim Prof As Variant
    Dim strProf As String
    Dim prenom As String, nom As String
    Dim profID As Long
    Dim total(0 To 4) As Currency
    Dim tauxHoraire As Currency
    Dim currentRow As Integer
    currentRow = 6
    
    Dim valeurTEC As Currency
    For i = 0 To 10 Step 5
        If i = 0 Then
            ws.Range("D" & currentRow).value = "EXCLUANT les clients '2000'"
        ElseIf i = 5 Then
            ws.Range("D" & currentRow).value = "SEULEMENT les clients '2000'"
        Else
            ws.Range("D" & currentRow).value = "TOUS LES CLIENTS"
        End If
        ws.Range("D" & currentRow).Font.Bold = True
        Erase total
        totalValeurTEC = 0
        currentRow = currentRow + 1
        
        For Each Prof In Fn_Sort_Dictionary_By_Keys(dictHours) 'Sort dictionary by keys in ascending order
            strProf = Mid(Prof, 4)
            profID = Fn_GetID_From_Initials(strProf)
            prenom = Fn_Get_Prenom_From_Initials(strProf)
            nom = Fn_Get_Nom_From_Initials(strProf)
            prenom = prenom & " " & Left(nom, 1) & "."
            tauxHoraire = Fn_Get_Hourly_Rate(profID, ws.Range("L3").value)
            ws.Range("E" & currentRow).value = prenom
            ws.Range("F" & currentRow).HorizontalAlignment = xlRight
            ws.Range("F" & currentRow).value = Format$(dictHours(Prof)(i + 0), "#,##0.00")
            ws.Range("G" & currentRow).HorizontalAlignment = xlRight
            ws.Range("G" & currentRow).value = Format$(tauxHoraire, "#,##0.00 $")
            ws.Range("H" & currentRow).HorizontalAlignment = xlRight
            ws.Range("H" & currentRow).value = Format$(dictHours(Prof)(i + 0) * tauxHoraire, "###,##0.00 $")
            ws.Range("I" & currentRow).HorizontalAlignment = xlRight
            ws.Range("I" & currentRow).value = Format$(dictHours(Prof)(i + 1), "#,##0.00")
            ws.Range("J" & currentRow).HorizontalAlignment = xlRight
            ws.Range("J" & currentRow).value = Format$(dictHours(Prof)(i + 2), "#,##0.00")
            ws.Range("K" & currentRow).HorizontalAlignment = xlRight
            ws.Range("K" & currentRow).value = Format$(dictHours(Prof)(i + 3), "#,##0.00")
            ws.Range("L" & currentRow).HorizontalAlignment = xlRight
            ws.Range("L" & currentRow).value = Format$(dictHours(Prof)(i + 4), "#,##0.00")
            currentRow = currentRow + 1
            
            Dim ii As Integer
            For ii = 0 To 4
                total(ii) = total(ii) + dictHours(Prof)(i + ii)
            Next ii
            
            totalValeurTEC = totalValeurTEC + (dictHours(Prof)(i) * tauxHoraire)
            If i = 0 Then
                valeurTEC = valeurTEC + (dictHours(Prof)(i) * tauxHoraire)
            End If

        Next Prof
        
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
                .color = 65535
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
    Dim solde As Double
    solde = Fn_Get_GL_Account_Balance("1210", maxDate)
    
    'Afficher le solde des TEC au Grand Livre
    ws.Range("D3").Font.Bold = True
    ws.Range("D3").Font.size = 12
    ws.Range("D3").Font.color = vbRed
    Dim message As String
    message = "Le solde au grand livre pour les TEC est de " & Format$(solde, "###,##0.00 $")
    If valeurTEC = solde Then
        message = message & ", donc aucune écriture"
    ElseIf valeurTEC > solde Then
        message = message & ", donc un Débit de " & Format$(valeurTEC - solde, "###,##0.00 $")
    Else
        message = message & ", donc un Crédit de " & Format$(valeurTEC - solde, "###,##0.00 $")
    End If
    ws.Range("D3").value = message
    
    ws.Shapes("Impression").Visible = True
    
    'Libérer la mémoire
    Set dictHours = Nothing
    Set Prof = Nothing
    Set ws = Nothing
    Set wsSource = Nothing
    
End Sub

Sub shp_TEC_Evaluation_Impression_Click()

    Call Evaluation_Apercu_Avant_Impression

End Sub

Sub Evaluation_Apercu_Avant_Impression()

    Dim ws As Worksheet: Set ws = wshTEC_Evaluation
    
    Dim rngToPrint As Range
    Set rngToPrint = ws.Range("C1:M31")
    
    Application.EnableEvents = False

    'Caractères pour le rapport
    With rngToPrint.Font
        .Name = "Aptos Narrow"
        .size = 10
    End With
    
    Application.EnableEvents = True
    
    DoEvents

    Dim header1 As String: header1 = "Évaluation des TEC au  " & wshTEC_Evaluation.Range("L3").value
    Dim header2 As String
    
    Call Simple_Print_Setup(wshTEC_Evaluation, rngToPrint, header1, header2, "$1:$1", "P")

    ws.PrintPreview
    
    'Libérer la mémoire
    Set rngToPrint = Nothing
    Set ws = Nothing
    
End Sub

Sub shp_TEC_Evaluation_Back_To_TEC_Menu_Click()

    Call TEC_Evaluation_Back_To_TEC_Menu
    
End Sub

Sub TEC_Evaluation_Back_To_TEC_Menu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshTEC_Evaluation:TEC_Evaluation_Back_To_TEC_Menu", 0)
    
    wshTEC_Evaluation.Visible = xlSheetVeryHidden
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select
    
    Call Log_Record("wshTEC_Evaluation:TEC_Evaluation_Back_To_TEC_Menu", startTime)

End Sub

