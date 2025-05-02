Attribute VB_Name = "modDev_Proof"
Option Explicit

Sub ObtenirHeuresFactur�esParFacture()

    '1. Obtenir toutes les charges factur�es
    Dim ws As Worksheet: Set ws = wsdTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim s As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 3 To lastUsedRow
        If ws.Cells(i, 16).value >= "24-24609" Then
            s = ws.Cells(i, 16).value & "-" & Format$(ws.Cells(i, 2), "00")
            'Ajoute au sommaire par facture / par ProfID
            If dict.Exists(s) Then
                dict(s) = dict(s) + ws.Cells(i, 8).value
            Else
                dict.Add s, ws.Cells(i, 8).value
            End If
        End If
    Next i

    'Cr�ation/Initialisation d'une feuille
    Dim feuilleNom As String
    feuilleNom = "X_Heures_Factur�es_Par_Facture"
    Call Erase_And_Create_Worksheet(feuilleNom)
    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Sheets(feuilleNom)
    wsOutput.Cells(1, 1).value = "Num�roFact"
    wsOutput.Cells(1, 2).value = "Prof"
    wsOutput.Cells(1, 3).value = "HeuresFact"
    
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
            wsOutput.Cells(r, 1).value = Left$(key, 8)
            wsOutput.Cells(r, 2).value = prof
            wsOutput.Cells(r, 3).value = dict(key)
            wsOutput.Cells(r, 3).NumberFormat = "##0.00"
        Next key
        Call SoustotalHeures(wsOutput, saveInvNo, r, st)
        
        r = r + 2
        wsOutput.Cells(r, 1).value = "* TOTAL *"
        wsOutput.Cells(r, 4).value = t
        wsOutput.Cells(r, 4).NumberFormat = "##0.00"
        wsOutput.Cells(r, 4).Font.Bold = True
        
    End If
    
End Sub

Sub SoustotalHeures(ws As Worksheet, saveInv As String, ByRef r As Long, ByRef st As Currency)

    If saveInv <> "" Then
        With ws.Cells(r, 3).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        r = r + 1
        ws.Cells(r, 4).value = st
        ws.Cells(r, 4).NumberFormat = "##0.00"
        st = 0
    End If

End Sub

Sub Identifier�cartsDeuxSourcesDeFacture() '2024-12-12 @ 10:55

    'Initialisation
    Dim wsEntete As Worksheet
    Set wsEntete = wsdFAC_Ent�te
    Dim wsComptesClients As Worksheet
    Set wsComptesClients = wsdFAC_Comptes_Clients
    Dim wsRapport As Worksheet
    
    On Error Resume Next
    Set wsRapport = ThisWorkbook.Worksheets("Rapport�cartsFactures")
    On Error GoTo 0
    If wsRapport Is Nothing Then
        Set wsRapport = ThisWorkbook.Worksheets.Add
        wsRapport.Name = "Rapport�cartsFactures"
    End If
    
    'Effacer le contenu du rapport
    wsRapport.Cells.Clear
    wsRapport.Cells(1, 1).value = "Num�ro de facture"
    wsRapport.Cells(1, 2).value = "$ FAC_Ent�te"
    wsRapport.Cells(1, 3).value = "$ FAC_Comptes_Clients"
    wsRapport.Cells(1, 4).value = "Diff�rence"
    
    'Charger les donn�es dans des dictionnaires
    Dim dictEntete As Object
    Set dictEntete = CreateObject("Scripting.Dictionary")
    Dim dictComptesClients As Object
    Set dictComptesClients = CreateObject("Scripting.Dictionary")
    
    'Lire wsdFAC_Ent�te
    Dim Facture As String
    Dim lastRowEntete As Long, lastRowComptes As Long
    lastRowEntete = wsEntete.Cells(wsEntete.Rows.count, 1).End(xlUp).row
    Dim montantEntete As Currency, totalEnt�teCC As Currency
    Dim i As Long
    For i = 3 To lastRowEntete
        Facture = wsEntete.Cells(i, fFacEInvNo).value
        montantEntete = wsEntete.Cells(i, fFacEARTotal).value
        totalEnt�teCC = totalEnt�teCC + montantEntete
        If Len(Facture) > 0 Then dictEntete(Facture) = montantEntete
    Next i
    
    'Lire wsdFAC_Comptes_Clients
    Dim montantCompte As Currency, totalComptesClients As Currency
    Dim montantPay� As Currency, montantR�gul As Currency
    Dim solde As Currency, soldeCC1 As Currency, soldeCC2 As Currency
    lastRowComptes = wsComptesClients.Cells(wsComptesClients.Rows.count, 1).End(xlUp).row
    For i = 3 To lastRowComptes
        Facture = wsComptesClients.Cells(i, fFacCCInvNo).value
        montantCompte = wsComptesClients.Cells(i, fFacCCTotal).value
        totalComptesClients = totalComptesClients + montantCompte
        montantPay� = wsComptesClients.Cells(i, fFacCCTotalPaid).value
        montantR�gul = wsComptesClients.Cells(i, fFacCCTotalRegul).value
        solde = wsComptesClients.Cells(i, fFacCCBalance).value
        If solde <> montantCompte - montantPay� - montantR�gul Then Stop
        soldeCC1 = soldeCC1 + solde
        soldeCC2 = soldeCC2 + montantCompte - montantPay� - montantR�gul
        If soldeCC1 <> soldeCC2 Then Stop
        If Len(Facture) > 0 Then dictComptesClients(Facture) = montantCompte
    Next i
    
    'Comparer et g�n�rer le rapport
    Dim fact As Variant
    Dim rowRapport As Long
    rowRapport = 2
    For Each fact In dictEntete.keys
        If dictComptesClients.Exists(fact) Then
            montantEntete = dictEntete(fact)
            montantCompte = dictComptesClients(fact)
            If montantEntete <> montantCompte Then
                wsRapport.Cells(rowRapport, 1).value = fact
                wsRapport.Cells(rowRapport, 2).value = montantEntete
                wsRapport.Cells(rowRapport, 3).value = montantCompte
                wsRapport.Cells(rowRapport, 4).value = montantEntete - montantCompte
                rowRapport = rowRapport + 1
            End If
        Else
            'Facture manquante dans wsdFAC_Comptes_Clients
            wsRapport.Cells(rowRapport, 1).value = fact
            wsRapport.Cells(rowRapport, 2).value = dictEntete(fact)
            wsRapport.Cells(rowRapport, 3).value = "Manquant"
            wsRapport.Cells(rowRapport, 4).value = "N/A"
            rowRapport = rowRapport + 1
        End If
    Next fact
    
    'V�rifier les factures manquantes dans wsdFAC_Ent�te
    For Each fact In dictComptesClients.keys
        If Not dictEntete.Exists(fact) Then
            wsRapport.Cells(rowRapport, 1).value = fact
            wsRapport.Cells(rowRapport, 2).value = "Manquant"
            wsRapport.Cells(rowRapport, 3).value = dictComptesClients(fact)
            wsRapport.Cells(rowRapport, 4).value = "N/A"
            rowRapport = rowRapport + 1
        End If
    Next fact
    
    wsRapport.Cells(rowRapport, 1).value = "Total des factures (selon FAC_Ent�te) est de " & Format$(totalEnt�teCC, "###,##0.00$")
    rowRapport = rowRapport + 1
    wsRapport.Cells(rowRapport, 1).value = "Total des factures (selon FAC_Comptes_Clients) est de " & Format$(totalComptesClients, "###,##0.00$")
    rowRapport = rowRapport + 1
    wsRapport.Cells(rowRapport, 1).value = "Solde des Comptes Clients (selon FAC_Comptes_Clients) est de " & Format$(soldeCC1, "###,##0.00$")
    
    ' Ajuster la mise en forme
    wsRapport.Columns.AutoFit
    
    MsgBox "La comparaison est termin�e. V�rifiez l'onglet 'Rapport�cartsFactures'.", vbInformation
    
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

