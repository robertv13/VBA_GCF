Attribute VB_Name = "modTEC_Evaluation"
Option Explicit

Sub TEC_Evaluation_Procedure(cutoffDate As String)

    Dim maxDate As Date
    Dim Y As Integer, m As Integer, d As Integer
    Y = year(cutoffDate)
    m = month(cutoffDate)
    d = day(cutoffDate)
    maxDate = DateSerial(Y, m, d)
    
    Dim ws As Worksheet: Set ws = wshTEC_Evaluation
    Dim wsSource As Worksheet: Set wsSource = wshTEC_Local
    
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Cells(ws.rows.count, "A").End(xlUp).rows
    
    'Transfère la table en mémoire (arr)
    Dim arr As Variant
    arr = wsSource.Range("A3:P" & lastUsedRow).value
    
    'Dictionaire pour accumuler les heures par professionnel
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim codeClient As String, profInit As String
    Dim hresNettes As Currency, hresNFact As Currency, hresFact As Currency, hresTEC As Currency
    Dim totalHresTEC As Currency, totalValeurTEC As Currency
    Dim tableau As Variant
    Dim trancheAge As String
    Dim ageTEC As Long
    For i = 1 To UBound(arr, 1) - 1
        hresNettes = 0
        hresNFact = 0
        hresFact = 0
        hresTEC = 0
        If CDate(arr(i, 4)) <= maxDate Then
            profInit = Format$(arr(i, 2), "000") & arr(i, 3)
            If UCase(arr(i, 14)) = "FAUX" Then
                hresNettes = arr(i, 8)
            Else
                hresNettes = 0
            End If
            codeClient = CStr(arr(i, 5))
            If UCase(arr(i, 10)) = "FAUX" Or Fn_Is_Client_Facturable(codeClient) = False Then
                hresNFact = hresNettes
            Else
                hresFact = hresNettes
            End If
            If (UCase(arr(i, 12)) = "FAUX" Or CDate(arr(i, 13)) > maxDate) And hresFact > 0 Then
                hresTEC = hresFact
            Else
                hresTEC = 0
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
                    dictHours.Add profInit, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
                End If
                tableau = dictHours(profInit) 'Obtenir le tableau a partir du dictionary
                
                'Heures pour ce TEC
                tableau(0) = tableau(0) + hresTEC
                
                'Accumule heures selon l'âge du TEC
                Select Case trancheAge
                    Case "- de 30 jours"
                        tableau(1) = tableau(1) + hresTEC
                    Case "31 @ 60 jours"
                        tableau(2) = tableau(2) + hresTEC
                    Case "61 @ 90 jours"
                        tableau(3) = tableau(3) + hresTEC
                    Case "+ de 90 jours"
                        tableau(4) = tableau(4) + hresTEC
                    End Select
                dictHours(profInit) = tableau 'Replacer le tableau dans le dictionnaire
            
            End If
        End If
    Next i
    
    Dim prof As Variant
    Dim strProf As String
    Dim prenom As String, nom As String
    Dim profID As Long
    Dim total(0 To 4) As Currency
    Dim tauxHoraire As Currency
    Dim currentRow As Integer
    currentRow = 6
    
    For Each prof In Fn_Sort_Dictionary_By_Keys(dictHours) 'Sort dictionary by keys in ascending order
        strProf = Mid(prof, 4)
        profID = Fn_GetID_From_Initials(strProf)
        prenom = Fn_Get_Prenom_From_Initials(strProf)
        nom = Fn_Get_Nom_From_Initials(strProf)
        prenom = prenom & " " & Left(nom, 1) & "."
        tauxHoraire = Fn_Get_Hourly_Rate(profID, ws.Range("K3").value)
        ws.Range("D" & currentRow).value = prenom
        ws.Range("E" & currentRow).HorizontalAlignment = xlRight
        ws.Range("E" & currentRow).value = Format$(dictHours(prof)(0), "#,##0.00")
        ws.Range("F" & currentRow).HorizontalAlignment = xlRight
        ws.Range("F" & currentRow).value = Format$(tauxHoraire, "#,##0.00 $")
        ws.Range("G" & currentRow).HorizontalAlignment = xlRight
        ws.Range("G" & currentRow).value = Format$(dictHours(prof)(0) * tauxHoraire, "###,##0.00 $")
        ws.Range("H" & currentRow).HorizontalAlignment = xlRight
        ws.Range("H" & currentRow).value = Format$(dictHours(prof)(1), "#,##0.00")
        ws.Range("I" & currentRow).HorizontalAlignment = xlRight
        ws.Range("I" & currentRow).value = Format$(dictHours(prof)(2), "#,##0.00")
        ws.Range("J" & currentRow).HorizontalAlignment = xlRight
        ws.Range("J" & currentRow).value = Format$(dictHours(prof)(3), "#,##0.00")
        ws.Range("K" & currentRow).HorizontalAlignment = xlRight
        ws.Range("K" & currentRow).value = Format$(dictHours(prof)(4), "#,##0.00")
        
        Dim ii As Integer
        For ii = 0 To 4
            total(ii) = total(ii) + dictHours(prof)(ii)
        Next ii
        totalValeurTEC = totalValeurTEC + (dictHours(prof)(0) * tauxHoraire)
        currentRow = currentRow + 1
    Next prof
    
    currentRow = currentRow + 1
    With ws.Range("d" & currentRow)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .value = "* Totaux *"
    End With
    With ws.Range("E" & currentRow)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .value = Format$(total(0), "#,##0.00")
    End With
    With ws.Range("G" & currentRow)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .value = Format$(totalValeurTEC, "###,##0.00 $")
    End With
    With ws.Range("H" & currentRow)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .value = Format$(total(1), "#,##0.00")
    End With
    With ws.Range("I" & currentRow)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .value = Format$(total(2), "#,##0.00")
    End With
    With ws.Range("J" & currentRow)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .value = Format$(total(3), "#,##0.00")
    End With
    With ws.Range("K" & currentRow)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .value = Format$(total(4), "#,##0.00")
    End With
        
    'Obtenir le solde d'ouverture & les transactions pour le compte TEC au Grand Livre
    Dim solde As Double
    solde = Fn_Get_GL_Account_Balance("1210", maxDate)
    
    'Afficher le solde des TEC au Grand Livre
    ws.Range("D3").Font.Bold = True
    ws.Range("D3").Font.size = 12
    ws.Range("D3").Font.Color = vbRed
    Dim message As String
    message = "Le solde au grand livre pour les TEC est de " & Format$(solde, "###,##0.00 $")
    If totalValeurTEC = solde Then
        message = message & ", donc aucune écriture"
    ElseIf totalValeurTEC > solde Then
        message = message & ", donc un Débit de " & Format$(totalValeurTEC - solde, "###,##0.00 $")
    Else
        message = message & ", donc un Crédit de " & Format$(totalValeurTEC - solde, "###,##0.00 $")
    End If
    ws.Range("D3").value = message
    
    'Libérer la mémoire
    Set dictHours = Nothing
    Set prof = Nothing
    Set ws = Nothing
    Set wsSource = Nothing
    
End Sub

