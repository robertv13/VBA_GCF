Attribute VB_Name = "modDev_Proof"
Option Explicit

Sub ObtenirHeuresFacturéesParFacture()

    '1. Obtenir toutes les charges facturées
    Dim ws As Worksheet: Set ws = wshTEC_Local
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

    'Création/Initialisation d'une feuille
    Dim feuilleNom As String
    feuilleNom = "X_Heures_Facturées_Par_Facture"
    Call Erase_And_Create_Worksheet(feuilleNom)
    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Sheets(feuilleNom)
    wsOutput.Cells(1, 1).value = "NuméroFact"
    wsOutput.Cells(1, 2).value = "Prof"
    wsOutput.Cells(1, 3).value = "HeuresFact"
    
    Dim key As Variant
    Dim prof As String, profID As Long, saveInvNo As String
    Dim t As Currency, st As Currency
    Dim r As Long: r = 1
    If dict.count <> 0 Then
        For Each key In Fn_Sort_Dictionary_By_Keys(dict, False) 'Sort dictionary by hours in ascending order
            profID = Mid(key, 10, Len(key) - 2)
            prof = Fn_Get_Prof_From_ProfID(profID)
            If Left(key, 8) <> saveInvNo Then
                Call SoustotalHeures(wsOutput, saveInvNo, r, st)
            End If
            t = t + dict(key)
            st = st + dict(key)
            saveInvNo = Left(key, 8)
            r = r + 1
            wsOutput.Cells(r, 1).value = Left(key, 8)
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

Sub IdentifierÉcartsDeuxSourcesDeFacture() '2024-12-12 @ 10:55

    'Initialisation
    Dim wsEntete As Worksheet
    Set wsEntete = wshFAC_Entête
    Dim wsComptesClients As Worksheet
    Set wsComptesClients = wshFAC_Comptes_Clients
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
    wsRapport.Cells(1, 1).value = "Numéro de facture"
    wsRapport.Cells(1, 2).value = "$ FAC_Entête"
    wsRapport.Cells(1, 3).value = "$ FAC_Comptes_Clients"
    wsRapport.Cells(1, 4).value = "Différence"
    
    'Charger les données dans des dictionnaires
    Dim dictEntete As Object
    Set dictEntete = CreateObject("Scripting.Dictionary")
    Dim dictComptesClients As Object
    Set dictComptesClients = CreateObject("Scripting.Dictionary")
    
    'Lire wshFAC_Entête
    Dim Facture As String
    Dim lastRowEntete As Long, lastRowComptes As Long
    lastRowEntete = wsEntete.Cells(wsEntete.Rows.count, 1).End(xlUp).row
    Dim montantEntete As Currency, totalEntêteCC As Currency
    Dim i As Long
    For i = 3 To lastRowEntete
        Facture = wsEntete.Cells(i, fFacEInvNo).value
        montantEntete = wsEntete.Cells(i, fFacEARTotal).value
        totalEntêteCC = totalEntêteCC + montantEntete
        If Len(Facture) > 0 Then dictEntete(Facture) = montantEntete
    Next i
    
    'Lire wshFAC_Comptes_Clients
    Dim montantCompte As Currency, totalComptesClients As Currency
    Dim montantPayé As Currency, montantRégul As Currency
    Dim solde As Currency, soldeCC1 As Currency, soldeCC2 As Currency
    lastRowComptes = wsComptesClients.Cells(wsComptesClients.Rows.count, 1).End(xlUp).row
    For i = 3 To lastRowComptes
        Facture = wsComptesClients.Cells(i, fFacCCInvNo).value
        montantCompte = wsComptesClients.Cells(i, fFacCCTotal).value
        totalComptesClients = totalComptesClients + montantCompte
        montantPayé = wsComptesClients.Cells(i, fFacCCTotalPaid).value
        montantRégul = wsComptesClients.Cells(i, fFacCCTotalRegul).value
        solde = wsComptesClients.Cells(i, fFacCCBalance).value
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
                wsRapport.Cells(rowRapport, 1).value = fact
                wsRapport.Cells(rowRapport, 2).value = montantEntete
                wsRapport.Cells(rowRapport, 3).value = montantCompte
                wsRapport.Cells(rowRapport, 4).value = montantEntete - montantCompte
                rowRapport = rowRapport + 1
            End If
        Else
            'Facture manquante dans wshFAC_Comptes_Clients
            wsRapport.Cells(rowRapport, 1).value = fact
            wsRapport.Cells(rowRapport, 2).value = dictEntete(fact)
            wsRapport.Cells(rowRapport, 3).value = "Manquant"
            wsRapport.Cells(rowRapport, 4).value = "N/A"
            rowRapport = rowRapport + 1
        End If
    Next fact
    
    'Vérifier les factures manquantes dans wshFAC_Entête
    For Each fact In dictComptesClients.keys
        If Not dictEntete.Exists(fact) Then
            wsRapport.Cells(rowRapport, 1).value = fact
            wsRapport.Cells(rowRapport, 2).value = "Manquant"
            wsRapport.Cells(rowRapport, 3).value = dictComptesClients(fact)
            wsRapport.Cells(rowRapport, 4).value = "N/A"
            rowRapport = rowRapport + 1
        End If
    Next fact
    
    wsRapport.Cells(rowRapport, 1).value = "Total des factures (selon FAC_Entête) est de " & Format$(totalEntêteCC, "###,##0.00$")
    rowRapport = rowRapport + 1
    wsRapport.Cells(rowRapport, 1).value = "Total des factures (selon FAC_Comptes_Clients) est de " & Format$(totalComptesClients, "###,##0.00$")
    rowRapport = rowRapport + 1
    wsRapport.Cells(rowRapport, 1).value = "Solde des Comptes Clients (selon FAC_Comptes_Clients) est de " & Format$(soldeCC1, "###,##0.00$")
    
    ' Ajuster la mise en forme
    wsRapport.Columns.AutoFit
    
    msgBox "La comparaison est terminée. Vérifiez l'onglet 'RapportÉcartsFactures'.", vbInformation
    
End Sub

Sub AnalyserFichiersLogSaisieHeures() '2024-12-15 @ 11:03

    'Définir le chemin du dossier principal des fichiers de PROD
    Dim cheminDossier As String
    cheminDossier = "C:\VBA\GC_FISCALITÉ\GCF_DataFiles\"

    'Initialiser FileSystemObject
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim dossier As Object
    Set dossier = FSO.GetFolder(cheminDossier)
    
    'Mettre en place le fichier de sortie
    Dim output As String
    output = "X_AnalyseTransTEC"
    Call CreateOrReplaceWorksheet(output)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(output)
    Dim r As Long: r = 1
    ws.Cells(r, 1) = "TECID"
    ws.Cells(r, 2) = "timeStamp"
    ws.Cells(r, 3) = "Opération"
    ws.Cells(r, 4) = "Heures"
    ws.Cells(r, 5) = "Prof"
    ws.Cells(r, 6) = "dateTEC"
    ws.Cells(r, 7) = "NoClient"
    ws.Cells(r, 8) = "NomClient"
    ws.Cells(r, 9) = "Description"
    ws.Cells(r, 10) = "Fichier"
    ws.Cells(r, 11) = "Ligne"
    ws.Range("A1").CurrentRegion.offset(1).Clear
    r = r + 1

    'Appeler la fonction récursive pour analyser tous les fichiers
    Call AnalyserDossier(dossier, FSO, ws, r)

    'Tri des informations
    If r > 2 Then
        With ws.Sort
            .SortFields.Clear
            'First sort On NoTEC
            .SortFields.Add key:=ws.Range("A2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On timeStamp
            .SortFields.Add key:=ws.Range("B2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On dateTEC
            .SortFields.Add key:=ws.Range("F2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange ws.Range("A2:K" & r)
            .Apply 'Apply Sort
         End With
    End If
    
    msgBox "L'analyse est terminée.", vbInformation
    
End Sub

Sub AnalyserDossier(dossier As Object, FSO As Object, ws As Worksheet, r As Long)

    'Parcourir tous les fichiers de dossier
    Dim fichier As Object
    Dim cheminFichier As String
    Dim compteurLigne As Long
    Dim compteurOccurences As Integer
    
    For Each fichier In dossier.Files
        If fichier.Name Like "LogSaisieHeures.*" Then
            cheminFichier = fichier.path
            Debug.Print "#077 - Analyse du fichier : " & cheminFichier

            'Initialiser le compteur pour ce fichier
            compteurLigne = 0

            'Ouvrir le fichier pour lecture seulement
            Dim fichierSource As Object
            Set fichierSource = FSO.OpenTextFile(cheminFichier, ForReading)

            'Parcourir les lignes du fichier
            Dim ligne As String, user As String, timeStamp As String, version As String, oper As String
            Dim noTEC As Long
            Dim prof As String
            Dim dateTEC As Date
            Dim noClient As String, nomClient As String, desc As String, comm As String
            Dim hres As Currency
            Dim isFACT As Boolean
            
            Do Until fichierSource.AtEndOfStream
                ligne = fichierSource.ReadLine
                compteurLigne = compteurLigne + 1

                'Compter les occurrences de " | " dans la ligne
                compteurOccurences = CompterOccurrences(ligne, "|")
                If compteurOccurences = 0 Or compteurOccurences = 4 Then
                    Exit Do
                End If
'                Debug.Print compteurOccurences & " - " & ligne
                Call FnStripLigneLogSaisieHeures(compteurOccurences, ligne, user, timeStamp, version, _
                                                 oper, noTEC, prof, dateTEC, noClient, nomClient, desc, _
                                                 hres, comm, isFACT)
                'Ajustement de certaies variables
                oper = UCase(Trim(oper))
                If InStr(oper, "ADD ") = 1 Then
                    noTEC = Mid(oper, InStr(oper, " ") + 1)
                    oper = Left(oper, InStr(oper, " ") - 1)
                End If
                If InStr(oper, "UPDATE ") = 1 Then
                    noTEC = Mid(oper, InStr(oper, " ") + 1)
                    oper = Left(oper, InStr(oper, " ") - 1)
                End If
                If InStr(oper, "DELETE-") = 1 Then
                    noTEC = Mid(oper, InStr(oper, "-") + 1)
                    oper = Left(oper, InStr(oper, "-") - 1)
                End If
                    
                ws.Cells(r, 1) = noTEC
                ws.Cells(r, 2) = timeStamp
                ws.Cells(r, 3) = oper
                ws.Cells(r, 4) = hres
                ws.Cells(r, 5) = prof
                ws.Cells(r, 6) = dateTEC
                ws.Cells(r, 7) = noClient
                ws.Cells(r, 8) = nomClient
                ws.Cells(r, 9) = desc
                ws.Cells(r, 10) = cheminFichier
                ws.Cells(r, 11) = compteurLigne
                r = r + 1
            Loop

            'Fermer le fichier qui a été lu
            fichierSource.Close
            
        End If
    Next fichier

    'Parcourir tous les sous-dossiers
    Dim sousDossier As Object
    For Each sousDossier In dossier.SubFolders
        Call AnalyserDossier(sousDossier, FSO, ws, r)
    Next sousDossier

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

Sub FnStripLigneLogSaisieHeures(nbrChamp As Integer, l As String, user As String, timeStamp As String, _
                                version As String, oper As String, noTEC As Long, prof As String, dateTEC As Date, _
                                noClient As String, nomClient As String, desc As String, hres As Currency, _
                                comm As String, isFACT As Boolean)

    Dim arr As Variant
    arr = Split(l, "|")
    
    Select Case nbrChamp
        Case 0
        
        Case 3
            user = Trim(arr(0))
            timeStamp = Trim(arr(1))
            If Len(timeStamp) <> 19 Then Stop
            oper = "Delete"
            noTEC = Abs(arr(3))
            prof = ""
            dateTEC = #7/31/2023#
            noClient = ""
            nomClient = ""
            desc = ""
            hres = 0
            isFACT = False
        Case 4
        
        Case 11
            If InStr(arr(0), ".") Then arr(0) = Left(arr(0), InStr(arr(0), ".") - 1)
            If IsDate(arr(0)) = False Then
                user = Trim(arr(0))
                timeStamp = Trim(arr(1))
                If Len(timeStamp) <> 19 And Len(timeStamp) < 22 Then Stop
                oper = Trim(arr(2))
                noTEC = Trim(arr(3))
                prof = Trim(arr(4))
                dateTEC = arr(5)
                noClient = Trim(arr(6))
                nomClient = Trim(arr(7))
                desc = Trim(arr(8))
                hres = Trim(Replace(arr(9), ".", ","))
                comm = Trim(arr(10))
                isFACT = arr(11)
            Else
                If InStr(Trim(arr(2)), "APP_") <> 1 Then
                    timeStamp = Trim(arr(0))
                    If Len(timeStamp) <> 19 And Len(timeStamp) < 23 Then Stop
                    user = Trim(arr(1))
                    oper = Trim(arr(2))
                    noTEC = Trim(arr(3))
                    prof = Trim(arr(4))
                    dateTEC = arr(5)
                    noClient = Trim(arr(6))
                    nomClient = Trim(arr(7))
                    desc = Trim(arr(8))
                    isFACT = arr(9)
                    hres = 0
                    comm = Trim(arr(10))
                Else
                    timeStamp = Trim(arr(0))
                    If Len(timeStamp) <> 19 And Len(timeStamp) < 23 Then Stop
                    user = Trim(arr(1))
                    version = Trim(arr(2))
                    oper = Trim(arr(3))
                    noTEC = 0
                    prof = Trim(arr(4))
                    dateTEC = arr(5)
                    noClient = Trim(arr(6))
                    nomClient = Trim(arr(7))
                    desc = Trim(arr(8))
                    hres = Trim(Replace(arr(9), ".", ","))
                    isFACT = Trim(arr(10))
                    comm = Trim(arr(11))
                End If
            End If
        Case 12
            If InStr(arr(0), ".") Then arr(0) = Left(arr(0), InStr(arr(0), ".") - 1)
            If IsDate(arr(0)) = False Then
                user = Trim(arr(0))
                timeStamp = Trim(arr(1))
                If Len(timeStamp) <> 19 And Len(timeStamp) < 23 Then Stop
                oper = Trim(arr(2))
                noTEC = Trim(arr(3))
                prof = Trim(arr(4))
                dateTEC = arr(5)
                noClient = Trim(arr(6))
                nomClient = Trim(arr(7))
                desc = Trim(arr(8))
                hres = Trim(Replace(arr(9), ".", ","))
                isFACT = arr(10)
                comm = Trim(arr(12))
            Else
                If IsDate(arr(4)) = False Then
                    If InStr(Trim(arr(2)), "APP_") <> 1 Then
                        timeStamp = Trim(arr(0))
                        If Len(timeStamp) <> 19 Then Stop
                        user = Trim(arr(1))
                        oper = Trim(arr(2))
                        noTEC = Trim(arr(3))
                        prof = Trim(arr(4))
                        dateTEC = arr(5)
                        noClient = Trim(arr(6))
                        nomClient = Trim(arr(7))
                        desc = Trim(arr(8))
                        hres = Trim(Replace(arr(9), ".", ","))
                        isFACT = Trim(arr(10))
                        comm = Trim(arr(12))
                    Else
                        timeStamp = Trim(arr(0))
                        If Len(timeStamp) <> 19 Then Stop
                        user = Trim(arr(1))
                        version = Trim(arr(2))
                        oper = Trim(arr(3))
                        noTEC = 0
                        prof = Trim(arr(4))
                        dateTEC = arr(5)
                        noClient = Trim(arr(6))
                        nomClient = Trim(arr(7))
                        desc = Trim(arr(8))
                        hres = Trim(Replace(arr(9), ".", ","))
                        isFACT = Trim(arr(10))
                        comm = Trim(arr(11))
                    End If
                Else
                    timeStamp = Trim(arr(0))
                    If Len(timeStamp) <> 19 Then Stop
                    user = Trim(arr(1))
                    oper = Trim(arr(2))
                    noTEC = 0
                    prof = Trim(arr(3))
                    dateTEC = arr(4)
                    noClient = Trim(arr(5))
                    nomClient = Trim(arr(6))
                    desc = Trim(arr(7))
                    hres = Trim(Replace(arr(8), ".", ","))
                    isFACT = Trim(arr(9))
                    comm = Trim(arr(10))
                End If
            End If
        Case 13
            If IsNumeric(arr(10)) = True Then
                timeStamp = Trim(arr(0))
                If Len(timeStamp) <> 19 And Len(timeStamp) < 23 Then Stop
                user = Trim(arr(1))
                oper = Trim(arr(3))
                noTEC = Trim(arr(4))
                prof = Trim(arr(5))
                dateTEC = arr(6)
                noClient = Trim(arr(7))
                nomClient = Trim(arr(8))
                desc = Trim(arr(9))
                hres = Trim(Replace(arr(10), ".", ","))
                isFACT = arr(11)
                comm = Trim(arr(12))
            Else
                timeStamp = Trim(arr(0))
                If Len(timeStamp) <> 19 And Len(timeStamp) < 23 Then Stop
                user = Trim(arr(1))
                oper = Trim(arr(2))
                noTEC = Trim(arr(3))
                prof = Trim(arr(4))
                dateTEC = arr(5)
                noClient = Trim(arr(6))
                nomClient = Trim(arr(7))
                desc = Trim(arr(8))
                hres = Trim(Replace(arr(9), ".", ","))
                isFACT = Trim(arr(10))
                comm = Trim(arr(11))
            End If
        Case Else
            Stop
    End Select
    
End Sub
