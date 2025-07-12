Attribute VB_Name = "modDev_Proof"
Option Explicit

Sub ObtenirHeuresFacturéesParFacture() '2025-04-07 @ 04:51

    '1. Obtenir toutes les charges facturées
    Dim ws As Worksheet: Set ws = wsdTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim s As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 3 To lastUsedRow
        If ws.Cells(i, 16).Value >= "24-24609" Then
            s = ws.Cells(i, 16).Value & "-" & Format$(ws.Cells(i, 2), "00")
            'Ajoute au sommaire par facture / par ProfID
            If dict.Exists(s) Then
                dict(s) = dict(s) + ws.Cells(i, 8).Value
            Else
                dict.Add s, ws.Cells(i, 8).Value
            End If
        End If
    Next i

    'Création/Initialisation d'une feuille
    Dim feuilleNom As String
    feuilleNom = "X_Heures_Facturées_Par_Facture"
    Call Erase_And_Create_Worksheet(feuilleNom)
    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Sheets(feuilleNom)
    wsOutput.Cells(1, 1).Value = "NuméroFact"
    wsOutput.Cells(1, 2).Value = "Prof"
    wsOutput.Cells(1, 3).Value = "HeuresFact"
    
    Dim key As Variant
    Dim prof As String, profID As Long, saveInvNo As String
    Dim t As Currency, st As Currency
    Dim r As Long: r = 1
    If dict.count <> 0 Then
        For Each key In Fn_Sort_Dictionary_By_Keys(dict, False) 'Sort dictionary by hours in ascending order
            profID = Mid$(key, 10, Len(key) - 2)
            prof = Fn_Get_Prof_From_ProfID(profID)
            If Left$(key, 8) <> saveInvNo Then
                Call SoustotalHeures(wsOutput, saveInvNo, r, st)
            End If
            t = t + dict(key)
            st = st + dict(key)
            saveInvNo = Left$(key, 8)
            r = r + 1
            wsOutput.Cells(r, 1).Value = Left$(key, 8)
            wsOutput.Cells(r, 2).Value = prof
            wsOutput.Cells(r, 3).Value = dict(key)
            wsOutput.Cells(r, 3).NumberFormat = "##0.00"
        Next key
        Call SoustotalHeures(wsOutput, saveInvNo, r, st)
        
        r = r + 2
        wsOutput.Cells(r, 1).Value = "* TOTAL *"
        wsOutput.Cells(r, 4).Value = t
        wsOutput.Cells(r, 4).NumberFormat = "##0.00"
        wsOutput.Cells(r, 4).Font.Bold = True
        
    End If
    
End Sub

Sub SoustotalHeures(ws As Worksheet, saveInv As String, ByRef r As Long, ByRef st As Currency)

    If saveInv <> vbNullString Then
        With ws.Cells(r, 3).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        r = r + 1
        ws.Cells(r, 4).Value = st
        ws.Cells(r, 4).NumberFormat = "##0.00"
        st = 0
    End If

End Sub

Sub IdentifierÉcartsDeuxSourcesDeFacture() '2024-12-12 @ 10:55

    'Initialisation
    Dim wsEntete As Worksheet
    Set wsEntete = wsdFAC_Entete
    Dim wsComptesClients As Worksheet
    Set wsComptesClients = wsdFAC_Comptes_Clients
    Dim wsRapport As Worksheet
    
    On Error Resume Next
    Set wsRapport = ThisWorkbook.Worksheets("RapportÉcartsFactures")
    On Error GoTo 0
    If wsRapport Is Nothing Then
        Set wsRapport = ThisWorkbook.Worksheets.Add
        wsRapport.Name = "RapportÉcartsFactures"
    End If
    
    'Effacer le contenu du rapport
    wsRapport.Cells.Clear
    wsRapport.Cells(1, 1).Value = "Numéro de facture"
    wsRapport.Cells(1, 2).Value = "$ FAC_Entête"
    wsRapport.Cells(1, 3).Value = "$ FAC_Comptes_Clients"
    wsRapport.Cells(1, 4).Value = "Différence"
    
    'Charger les données dans des dictionnaires
    Dim dictEntete As Object
    Set dictEntete = CreateObject("Scripting.Dictionary")
    Dim dictComptesClients As Object
    Set dictComptesClients = CreateObject("Scripting.Dictionary")
    
    'Lire wsdFAC_Entete
    Dim Facture As String
    Dim lastRowEntete As Long, lastRowComptes As Long
    lastRowEntete = wsEntete.Cells(wsEntete.Rows.count, 1).End(xlUp).Row
    Dim montantEntete As Currency, totalEntêteCC As Currency
    Dim i As Long
    For i = 3 To lastRowEntete
        Facture = wsEntete.Cells(i, fFacEInvNo).Value
        montantEntete = wsEntete.Cells(i, fFacEARTotal).Value
        totalEntêteCC = totalEntêteCC + montantEntete
        If Len(Facture) > 0 Then dictEntete(Facture) = montantEntete
    Next i
    
    'Lire wsdFAC_Comptes_Clients
    Dim montantCompte As Currency, totalComptesClients As Currency
    Dim montantPayé As Currency, montantRégul As Currency
    Dim solde As Currency, soldeCC1 As Currency, soldeCC2 As Currency
    lastRowComptes = wsComptesClients.Cells(wsComptesClients.Rows.count, 1).End(xlUp).Row
    For i = 3 To lastRowComptes
        Facture = wsComptesClients.Cells(i, fFacCCInvNo).Value
        montantCompte = wsComptesClients.Cells(i, fFacCCTotal).Value
        totalComptesClients = totalComptesClients + montantCompte
        montantPayé = wsComptesClients.Cells(i, fFacCCTotalPaid).Value
        montantRégul = wsComptesClients.Cells(i, fFacCCTotalRegul).Value
        solde = wsComptesClients.Cells(i, fFacCCBalance).Value
        If solde <> montantCompte - montantPayé - montantRégul Then Stop
        soldeCC1 = soldeCC1 + solde
        soldeCC2 = soldeCC2 + montantCompte - montantPayé - montantRégul
        If soldeCC1 <> soldeCC2 Then Stop
        If Len(Facture) > 0 Then dictComptesClients(Facture) = montantCompte
    Next i
    
    'Comparer et générer le rapport
    Dim fact As Variant
    Dim rowRapport As Long
    rowRapport = 2
    For Each fact In dictEntete.keys
        If dictComptesClients.Exists(fact) Then
            montantEntete = dictEntete(fact)
            montantCompte = dictComptesClients(fact)
            If montantEntete <> montantCompte Then
                wsRapport.Cells(rowRapport, 1).Value = fact
                wsRapport.Cells(rowRapport, 2).Value = montantEntete
                wsRapport.Cells(rowRapport, 3).Value = montantCompte
                wsRapport.Cells(rowRapport, 4).Value = montantEntete - montantCompte
                rowRapport = rowRapport + 1
            End If
        Else
            'Facture manquante dans wsdFAC_Comptes_Clients
            wsRapport.Cells(rowRapport, 1).Value = fact
            wsRapport.Cells(rowRapport, 2).Value = dictEntete(fact)
            wsRapport.Cells(rowRapport, 3).Value = "Manquant"
            wsRapport.Cells(rowRapport, 4).Value = "N/A"
            rowRapport = rowRapport + 1
        End If
    Next fact
    
    'Vérifier les factures manquantes dans wsdFAC_Entete
    For Each fact In dictComptesClients.keys
        If Not dictEntete.Exists(fact) Then
            wsRapport.Cells(rowRapport, 1).Value = fact
            wsRapport.Cells(rowRapport, 2).Value = "Manquant"
            wsRapport.Cells(rowRapport, 3).Value = dictComptesClients(fact)
            wsRapport.Cells(rowRapport, 4).Value = "N/A"
            rowRapport = rowRapport + 1
        End If
    Next fact
    
    wsRapport.Cells(rowRapport, 1).Value = "Total des factures (selon FAC_Entête) est de " & Format$(totalEntêteCC, "###,##0.00$")
    rowRapport = rowRapport + 1
    wsRapport.Cells(rowRapport, 1).Value = "Total des factures (selon FAC_Comptes_Clients) est de " & Format$(totalComptesClients, "###,##0.00$")
    rowRapport = rowRapport + 1
    wsRapport.Cells(rowRapport, 1).Value = "Solde des Comptes Clients (selon FAC_Comptes_Clients) est de " & Format$(soldeCC1, "###,##0.00$")
    
    ' Ajuster la mise en forme
    wsRapport.Columns.AutoFit
    
    MsgBox "La comparaison est terminée. Vérifiez l'onglet 'RapportÉcartsFactures'.", vbInformation
    
End Sub

Function CompterOccurrences(texte As String, motif As String) As Long

    'Initialiser la position et trouver la position
    Dim position As Long
    position = InStr(1, texte, motif, vbTextCompare)

    'Parcourir le texte pour trouver toutes les occurrences
    Dim compteur As Long
    Do While position > 0
        compteur = compteur + 1
        position = InStr(position + Len(motif), texte, motif, vbTextCompare)
    Loop

    'Retourner le nombre d'occurrences
    CompterOccurrences = compteur
    
End Function

