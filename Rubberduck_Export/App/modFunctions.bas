Attribute VB_Name = "modFunctions"
Option Explicit

Function Fn_ProfIDAPartirDesInitiales(i As String)

    Dim cell As Range
    
    For Each cell In wsdADMIN.Range("dnrProf_All")
        If cell.Value2 = i Then
            Fn_ProfIDAPartirDesInitiales = cell.offset(0, 1).Value
            Exit Function
        End If
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Function Fn_InitialesAPartirProfID(i As Long)

    Dim cell As Range
    
    For Each cell In wsdADMIN.Range("dnrProf_All")
        If cell.Value2 = i Then
            Fn_InitialesAPartirProfID = cell.offset(0, -1).Value
            Exit Function
        End If
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Function Fn_ObtenirLigneDeFeuille(feuille As String, cle As Variant, cleCol As Integer) As Variant

    'Feuille à rechercher
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(feuille)
    
    'Charger les données en mémoire
    Dim allData As Variant
    allData = ws.usedRange.Value

    'Parcourir les données pour trouver la valeur
    Dim resultArray() As Variant
    Dim i As Long
    For i = 1 To UBound(allData, 1)
'        If i = 761 Or allData(i, 2) = "1472" Then Stop
        If Trim(allData(i, cleCol)) = Trim(cle) Then
            'Ligne est trouvée alors on copie toutes les colonnes dans le tableau résultat
            resultArray = Application.index(allData, i, 0)
            Fn_ObtenirLigneDeFeuille = resultArray
            Exit Function
        End If
    Next i
    
    'Si aucune correspondance n'a été trouvée, retourner une valeur vide
    Fn_ObtenirLigneDeFeuille = CVErr(xlErrValue)
    
End Function

Function Fn_ClientIDAPartirDuNomDeClient(nomClient As String) '2024-02-14 @ 06:07

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_ClientIDAPartirDuNomDeClient", nomClient, 0)
    
    Dim ws As Worksheet: Set ws = wsdBD_Clients
    
    On Error Resume Next
    Dim dynamicRange As Range: Set dynamicRange = ws.Range("dnrClients_All")
    On Error GoTo 0

    If ws Is Nothing Or dynamicRange Is Nothing Then
        MsgBox "La feuille 'Clients' ou le DynamicRange 'dnrClients_All' n'a pas été trouvé!", _
            vbExclamation
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly, reuires EXACT match (5th parameter = 0) - 2025-01-12 @ 14:48
    Dim result As Variant
    result = Application.WorksheetFunction.XLookup(nomClient, _
                                                   dynamicRange.Columns(1), _
                                                   dynamicRange.Columns(2), _
                                                   "Not Found", _
                                                   0, _
                                                   1)
    If result <> "Not Found" Then
        Fn_ClientIDAPartirDuNomDeClient = result
        ufSaisieHeures.txtClientID.Value = result
    Else
        MsgBox "Impossible de retrouver le nom du client dans la feuille" & vbNewLine & vbNewLine & _
               "BD_Clients (" & nomClient & ")" & vbNewLine & vbNewLine & _
               "VOUS DEVEZ SAISIR LE NOM DU CLIENT À NOUVEAU", _
               vbCritical, "Client INEXISTANT dans la base de données des clients"
    End If
    
    'Libérer la mémoire
    Set dynamicRange = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_ClientIDAPartirDuNomDeClient", vbNullString, startTime)

End Function

Function Fn_CellSpecifiqueDeBDClient(nomClient As String, ByRef colNumberSearch As Integer, ByRef colNumberData As Integer) As String '2025-10-31 @ 05:37

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_CellSpecifiqueDeBDClient", nomClient, 0)
    
    nomClient = Trim(nomClient)
    Dim ws As Worksheet: Set ws = wsdBD_Clients
    If ws Is Nothing Then
        MsgBox "La feuille 'Clients' est introuvable !", vbCritical
        Exit Function
    End If
    
    Dim dynamicRange As Range
    On Error Resume Next
    Set dynamicRange = ws.Range("dnrClients_All")
    On Error GoTo 0
    
    If dynamicRange Is Nothing Then
        MsgBox "Le DynamicRange 'dnrClients_All' n'a pas été trouvé!", _
            vbCritical, _
            "Problème important avec l'application"
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly, requires EXACT match (5th parameter = 0 ) - 2025-01-12 @ 14:49
    If colNumberSearch < 1 Or colNumberSearch > dynamicRange.Columns.count Then Exit Function
    If colNumberData < 1 Or colNumberData > dynamicRange.Columns.count Then Exit Function
    Dim result As Variant
    result = Application.XLookup(nomClient, _
                dynamicRange.Columns(colNumberSearch), _
                dynamicRange.Columns(colNumberData), _
                "Not Found", 0, 1)
'    result = Application.WorksheetFunction.XLookup(nomClient, _
'                                                   dynamicRange.Columns(colNumberSearch), _
'                                                   dynamicRange.Columns(colNumberData), _
'                                                   "Not Found", _
'                                                   0, _
'                                                   1)
    If Not result = "Not Found" Then
        Fn_CellSpecifiqueDeBDClient = result
    Else
        MsgBox _
            Prompt:="Impossible de retrouver ce nom du client dans BD_Clients" & vbNewLine & vbNewLine & _
                    "VOUS DEVEZ SAISIR À NOUVEAU LE NOM DU CLIENT", _
            Title:="Erreur grave - Impossible de retrouver le nom du client - ", _
            Buttons:=vbCritical
    End If
    
    Set dynamicRange = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_CellSpecifiqueDeBDClient", vbNullString, startTime)

End Function

Function Fn_ClientIDAPartirDuNomDeFournisseur(nomFournisseur As String) '2024-07-03 @ 16:13

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_ClientIDAPartirDuNomDeFournisseur", nomFournisseur, 0)
    
    Dim ws As Worksheet: Set ws = wsdBD_Fournisseurs
    
    On Error Resume Next
    Dim dynamicRange As Range: Set dynamicRange = ws.Range("dnrSuppliers_All")
    On Error GoTo 0

    If ws Is Nothing Or dynamicRange Is Nothing Then
        MsgBox "La feuille 'BD_Fournisseurs' ou le DynamicRange 'dnrSuppliers_All' n'a pas été trouvé!", _
            vbExclamation
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly
    Dim result As Variant
    result = Application.WorksheetFunction.XLookup(nomFournisseur, _
                                                   dynamicRange.Columns(1), _
                                                   dynamicRange.Columns(2), _
                                                   "Not Found", _
                                                   0, _
                                                   1)
    If result <> "Not Found" Then
        Fn_ClientIDAPartirDuNomDeFournisseur = result
    Else
        Fn_ClientIDAPartirDuNomDeFournisseur = 0
    End If
    
    'Libérer la mémoire
    Set dynamicRange = Nothing
    Set ws = Nothing

    Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_ClientIDAPartirDuNomDeFournisseur", vbNullString, startTime)

End Function

Function Fn_PrenomAPartirDesInitiales(i As String)

    Dim cell As Range
    
    For Each cell In wsdADMIN.Range("dnrProf_All")
        If cell.Value2 = i Then
            Fn_PrenomAPartirDesInitiales = cell.offset(0, 2).Value
            Exit Function
        End If
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Function Fn_NomAPartirDesInitiales(i As String)

    Dim cell As Range
    
    For Each cell In wsdADMIN.Range("dnrProf_All")
        If cell.Value2 = i Then
            Fn_NomAPartirDesInitiales = cell.offset(0, 3).Value
            Exit Function
        End If
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Function Fn_ValeurAPartirUniqueID(ws As Worksheet, uniqueID As String, keyColumn As Integer, returnColumn As Integer) As Variant

    'Définir la dernière ligne utilisée de la feuille
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, keyColumn).End(xlUp).Row
    
    'Définir la plage de recherche (toute la colonne de la clé)
    Dim searchRange As Range
    Set searchRange = ws.Range(ws.Cells(1, keyColumn), ws.Cells(lastRow, keyColumn))
    
    'Rechercher la clé dans la colonne spécifiée
    Dim foundCell As Range
    Set foundCell = searchRange.Find(What:=uniqueID, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Si on a trouvé 'uniqueID', retourner la valeur de la colonne de retour
    If Not foundCell Is Nothing Then
        Fn_ValeurAPartirUniqueID = ws.Cells(foundCell.row, returnColumn).Value
    Else
        'Si l'on a pas trouvée, retourner une valeur d'erreur ou un message
        Fn_ValeurAPartirUniqueID = "uniqueID introuvable"
    End If
    
    'Libérer la mémoire
    Set foundCell = Nothing
    Set searchRange = Nothing
    Set ws = Nothing
    
End Function

Function Fn_TrouveDataDansUnePlage(r As Range, cs As Long, ss As String, cr As Long) As Variant() '2024-03-29 @ 05:39
    
    'This function is used to retrieve information from in a range(r) at column (cs) the value of (ss)
    'If found, it returns an array, with the cell address(1), the row(2) and the value of column cr(3)
    'Otherwise it return an empty array
    '2024-03-09 - First version
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_TrouveDataDansUnePlage", vbNullString, 0)
    
    Dim foundInfo(1 To 3) As Variant 'Cell Address, Row, Value
    
    'Search for the string in a given range (r) at the column specified (cs)
    Dim foundCell As Range: Set foundCell = r.Columns(cs).Find(What:=ss, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the string was found
    If Not foundCell Is Nothing Then
        'With the foundCell get the the address, the row number and the value
        foundInfo(1) = foundCell.Address
        foundInfo(2) = foundCell.row
        foundInfo(3) = foundCell.offset(0, cr - cs).Value 'Return Column - Searching column
        Fn_TrouveDataDansUnePlage = foundInfo 'foundInfo is an array
    Else
        Fn_TrouveDataDansUnePlage = foundInfo 'foundInfo is an array
    End If
    
    'Libérer la mémoire
    Set foundCell = Nothing

    Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_TrouveDataDansUnePlage", vbNullString, startTime)

End Function

Function Fn_DetruireLigneSiValeurEstTrouvee(valueToFind As Variant, hono As Double) As String '2025-07-11 @ 01:12
    
    'Define the worksheet
    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Details
    
    Dim rowsToDelete As Collection: Set rowsToDelete = New Collection
    Dim lastUsedRow As Long
    Dim i As Long
    
    With ws
        lastUsedRow = .Cells(.Rows.count, 1).End(xlUp).Row
        For i = 2 To lastUsedRow
            If .Cells(i, 2).Value = valueToFind And (UCase(.Cells(i, 9).Value) = "FAUX" Or .Cells(i, 9).Value = 0) Then
                rowsToDelete.Add i
            End If
        Next i
    End With
    
'    'Define the range to search in (Column 1)
'    Dim searchRange As Range: Set searchRange = ws.Columns(2)
'
'    'Search for the first occurrence of the value
'    Dim cell As Range
'    Set cell = searchRange.Find(What:=valueToFind, _
'                                LookIn:=xlValues, _
'                                LookAt:=xlWhole)
'
'    'Check if the value is found
'    Dim firstAddress As String

    If rowsToDelete.count > 0 Then
'    If Not cell Is Nothing Then
'        Fn_DetruireLigneSiValeurEstTrouvee = firstAddress
        
        'Confirm with the user
        Dim reponse As Long
        reponse = MsgBox("Il existe déjà une demande de facture pour ce client" & _
                  vbNewLine & "au montant de " & Format$(hono, "#,##0.00$") & _
                  vbNewLine & vbNewLine & "Désirez-vous..." & vbNewLine & vbNewLine & _
                  "   1) (OUI) REMPLACER cette demande" & vbNewLine & vbNewLine & _
                  "   2) (NON) pour NE RIEN CHANGER à la demande existante" & vbNewLine & vbNewLine & _
                  "   3) (ANNULER) pour ANNULER la demande", vbYesNoCancel, "Confirmation pour un projet existant")
        Select Case reponse
            Case vbYes, vbCancel
                If reponse = vbYes Then
                    Fn_DetruireLigneSiValeurEstTrouvee = "REMPLACER"
                End If
                If reponse = vbCancel Then
                    Fn_DetruireLigneSiValeurEstTrouvee = "SUPPRIMER"
                End If
                
                'Soft delete all collected rows from wsdFAC_Projets_Details (locally) - 2025-07-11 @ 00:58
                Dim lo As ListObject
                Set lo = ws.ListObjects("l_tbl_FAC_Projets_Details")
                
                For i = rowsToDelete.count To 1 Step -1
                    ws.Cells(rowsToDelete(i), 9).Value = -1
                Next i
                
                'Soft Delete FAC_Projets_Entete
                Set lo = wsdFAC_Projets_Entete.ListObjects("l_tbl_FAC_Projets_Entete")
                
                For i = 1 To lo.ListRows.count
                    If lo.ListRows(i).Range.Cells(1, 2).Value = valueToFind Then
                        lo.ListRows(i).Range(1, 26).Value = -1
                    End If
                Next i
                
'                Set cell = lo.ListColumns(2).DataBodyRange.Find(valueToFind, LookIn:=xlValues, LookAt:=xlWhole)
'
'                If Not cell Is Nothing Then
'                    lo.ListRows(cell.row).Range.Cells(1, 26).Value = -1
'                Else
'                    MsgBox "Valeur non trouvée"
'                End If
'
                'Update rows from MASTER file (details)
                Dim destinationFileName As String, destinationTab As String
                destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                                      wsdADMIN.Range("MASTER_FILE").Value
                destinationTab = "FAC_Projets_Details$"
                
                Dim columnName As String
                columnName = "NomClient"
                Call DetruireDetailSiEnteteEstDetruite(destinationFileName, _
                                                        destinationTab, _
                                                        columnName, _
                                                        valueToFind)
                                                                     
                'Update row from MASTER file (entête)
                destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                                      wsdADMIN.Range("MASTER_FILE").Value
                destinationTab = "FAC_Projets_Entete$"
                Call DetruireEnteteSiEnteteEstDetruite(destinationFileName, _
                                                        destinationTab, _
                                                        columnName, _
                                                        valueToFind) '2024-07-19 @ 15:31
            Case vbNo
                Fn_DetruireLigneSiValeurEstTrouvee = "RIEN_CHANGER"
        End Select
    Else
        Fn_DetruireLigneSiValeurEstTrouvee = "REMPLACER"
    End If
    
    'Libérer la mémoire
    Set lo = Nothing
    Set rowsToDelete = Nothing
    Set ws = Nothing
    
End Function

Function Fn_GetCheckBoxPosition(chkBox As OLEObject) As String

    'Get the cell that contains the top-left corner of the CheckBox
    Fn_GetCheckBoxPosition = chkBox.TopLeftCell.Address
    
End Function

Function Fn_TypeDonneeColonne(col As Range) As String

    Dim cell As Range
    Dim dataType As String
    Dim cellValue As Variant
    
    dataType = "Empty" ' Default type if no data found
    
    ' Loop through cells in the first few rows to determine the data type
    For Each cell In col.Cells
        cellValue = cell.Value
        If Not IsEmpty(cellValue) Then
            If IsNumeric(cellValue) Then
                If IsDate(cellValue) Then
                    dataType = "Date"
                Else
                    dataType = "Numeric"
                End If
            ElseIf IsDate(cellValue) Then
                dataType = "Date"
            ElseIf IsError(cellValue) Then
                dataType = "Error"
            Else
                Select Case VarType(cellValue)
                    Case vbString
                        dataType = "Text"
                    Case vbBoolean
                        dataType = "Boolean"
                    Case vbDate
                        dataType = "Date"
                    Case Else
                        dataType = "Other"
                End Select
            End If
            ' Exit loop once a non-empty value is found
            Exit For
        End If
    Next cell
    
    Fn_TypeDonneeColonne = dataType
    
    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Public Function Fn_GetGL_Code_From_GL_Description(glDescr As String) 'XLOOKUP - 2024-01-09 @ 09:19

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_GetGL_Code_From_GL_Description", glDescr, 0)
    
    Dim ws As Worksheet: Set ws = wsdADMIN
    
    On Error Resume Next
    Dim dynamicRange As Range: Set dynamicRange = ws.Range("dnrPlanComptable_All")
    On Error GoTo 0
    
    If ws Is Nothing Or dynamicRange Is Nothing Then
        MsgBox "La feuille 'Admin' ou le DynamicRange n'a pas été trouvé!", _
            vbExclamation
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly
    Dim result As Variant
    result = Application.WorksheetFunction.XLookup(glDescr, _
        dynamicRange.Columns(1), dynamicRange.Columns(2), _
        "Not Found", 0, 1)
    
    Call modDev_Utils.EnregistrerLogApplication("     modFunctions:Fn_GetGL_Code_From_GL_Description - " & result, -1)
    
    If result <> "Not Found" Then
        Fn_GetGL_Code_From_GL_Description = result
    Else
        MsgBox "Impossible de retrouver la valeur dans la première colonne", vbExclamation
    End If

    'Libérer la mémoire
    Set dynamicRange = Nothing
    Set ws = Nothing

    Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_GetGL_Code_From_GL_Description", vbNullString, startTime)

End Function

Function Fn_DescriptionAPartirNoCompte(codeCompte As String) As String '2025-07-20 @ 08:11

    Dim planComptable As Variant
    Dim nbCol As Long
    nbCol = 2
    planComptable = Fn_PlanComptableTableau2D(nbCol)
    
    Fn_DescriptionAPartirNoCompte = "Compte non-défini"
   
    'Boucle à travers le tableau 'planComptable'
    Dim i As Long
    For i = LBound(planComptable) To UBound(planComptable)
        If Trim(planComptable(i, 1)) = Trim(codeCompte) Then
            Fn_DescriptionAPartirNoCompte = planComptable(i, 2)
            Exit Function
        End If
    Next i
    
End Function

Function Fn_TotalTransGLMois(glCode As String, dateFinMois As Date) As Double '2025-02-07 @ 13:46
    
    Fn_TotalTransGLMois = 0
    
    Dim dateDebutMois As Date
    dateDebutMois = DateSerial(year(dateFinMois), month(dateFinMois), 1)
    
    'AdvancedFilter GL_Trans with FromDate to ToDate, returns rngResult
    Dim rngResult As Range
    Call modGL_Stuff.ObtenirSoldeCompteEntreDebutEtFin(glCode, dateDebutMois, dateFinMois, rngResult)
    
    'Méthode plus rapide pour obtenir une somme
    Fn_TotalTransGLMois = Application.WorksheetFunction.Sum(rngResult.Columns(7)) _
                                           - Application.WorksheetFunction.Sum(rngResult.Columns(8))

End Function

Function Fn_TECTotalOuHeuresPourFactureAvecAF(invNo As String, t As String) As Currency

    'Le type (t) est "Heures" -OU- "Dollars", selon le type le total des Heures ou des Dollars
    
    Fn_TECTotalOuHeuresPourFactureAvecAF = 0
    
    Dim ws As Worksheet: Set ws = wsdFAC_Details
    
    'Effacer les données de la dernière utilisation
    ws.Range("I6:I10").ClearContents
    ws.Range("I6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_FAC_Details[#All]")
    ws.Range("I7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("I2:I3")
    ws.Range("I3").Value = invNo
    ws.Range("I8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("K1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("K2:N2")
    ws.Range("I9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "K").End(xlUp).Row
    ws.Range("I10").Value = lastUsedRow - 2 & " lignes"
    
    'Aucun tri nécessaire (besoins)
    If lastUsedRow > 2 Then
        Dim i As Long
        For i = 3 To lastUsedRow
            If InStr(ws.Cells(i, 11), "*** - [Sommaire des TEC] pour la facture - ") = 1 Then
                If t = "Heures" Then
                    Fn_TECTotalOuHeuresPourFactureAvecAF = Fn_TECTotalOuHeuresPourFactureAvecAF + ws.Cells(i, "L")
                Else
                    Fn_TECTotalOuHeuresPourFactureAvecAF = Fn_TECTotalOuHeuresPourFactureAvecAF + ws.Cells(i, "N")
                End If
            End If
        Next i
    End If
    
    'Force un arrondissement à 2 décimales
    Fn_TECTotalOuHeuresPourFactureAvecAF = Round(Fn_TECTotalOuHeuresPourFactureAvecAF, 2)
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
End Function

Public Function Fn_Find_Row_Number_TECID(ByVal uniqueID As Variant, ByVal lookupRange As Range) As Long '2024-08-10 @ 05:41
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_Find_Row_Number_TECID", vbNullString, 0)
    
    On Error Resume Next
        Dim cell As Range
        Set cell = lookupRange.Find(What:=uniqueID, LookIn:=xlValues, LookAt:=xlWhole)
        If Not cell Is Nothing Then
            Fn_Find_Row_Number_TECID = cell.row
        Else
            Fn_Find_Row_Number_TECID = -1 'Not found
        End If
    On Error GoTo 0
    
    'Libérer la mémoire
    Set cell = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_Find_Row_Number_TECID", vbNullString, startTime)
    
End Function

Function Fn_PaiementsTotalPourFactureAvecAF(invNo As String)

    Fn_PaiementsTotalPourFactureAvecAF = 0
    
    Dim ws As Worksheet: Set ws = wsdENC_Details
    
    'Effacer les données de la dernière utilisation
    ws.Range("H6:H10").ClearContents
    ws.Range("H6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_ENC_Details[#All]")
    ws.Range("H7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("H2:H3")
    ws.Range("H3").Value = invNo
    ws.Range("H8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("J1").CurrentRegion
    rngResult.offset(3, 0).Clear
    Set rngResult = ws.Range("J3:N3")
    ws.Range("H9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
    ws.Range("H10").Value = lastUsedRow - 3 & " lignes"
    
    'Il n'est pas nécessaire de trier les résultats
    If lastUsedRow > 3 Then
        Set rngResult = ws.Range("J4:N" & lastUsedRow)
        Fn_PaiementsTotalPourFactureAvecAF = Application.WorksheetFunction.Sum(rngResult.Columns(5))
    End If

    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
End Function

Function Fn_RegularisationsTotalPourFactureAvecAF(invNo As String)

    Fn_RegularisationsTotalPourFactureAvecAF = 0
    
    Dim ws As Worksheet: Set ws = wsdCC_Regularisations
    
    'Effacer les données de la dernière utilisation
    ws.Range("M6:M10").ClearContents
    ws.Range("M6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_CC_Regularisations[#All]")
    ws.Range("M7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("M2:M3")
    ws.Range("M3").Value = invNo
    ws.Range("M8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("O1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("O2:Y2")
    ws.Range("M9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "O").End(xlUp).Row
    ws.Range("M10").Value = lastUsedRow - 2 & " lignes"
    
    'Il n'est pas nécessaire de trier les résultats
    Dim sommeRegul As Currency
    Dim i As Long
    If lastUsedRow > 2 Then
        Set rngResult = ws.Range("O3:Y" & lastUsedRow)
        For i = 5 To 9
            sommeRegul = sommeRegul + Application.WorksheetFunction.Sum(rngResult.Columns(i))
        Next i
    End If

    Fn_RegularisationsTotalPourFactureAvecAF = sommeRegul
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
End Function

Function Fn_CellAPartirUneFeuille(feuille As String, cle As String, cleCol As Integer, retourCol As Integer) As String '2025-03-04 @ 06:56

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_CellAPartirUneFeuille", cle, 0)
    
    'Définir la feuille pour la recherche
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(feuille)
    
    'Définir la plage pour le résultat
    Dim resultat As Range
    
    'Utilisation de la méthode Find pour rechercher dans la première colonne
    Set resultat = ws.Columns(cleCol).Find(What:=cle, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not resultat Is Nothing Then
        Fn_CellAPartirUneFeuille = ws.Cells(resultat.row, retourCol)
    Else
        Fn_CellAPartirUneFeuille = vbNullString
    End If
    
    'Libérer la mémoire
    Set resultat = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFunctions:Fn_CellAPartirUneFeuille", vbNullString, startTime)

End Function

Function Fn_ClientEstValide(clientCode As String) As Boolean '2024-10-26 @ 18:30

    '2024-08-14 @ 10:17 - Verify that a client exists, based on clientCode
    
    Fn_ClientEstValide = False
    
    Dim lastUsedRow As Long
    lastUsedRow = wsdBD_Clients.Cells(wsdBD_Clients.Rows.count, "B").End(xlUp).Row
    Dim rngToSearch As Range
    Set rngToSearch = wsdBD_Clients.Range("B1:B" & lastUsedRow)
    
    'Search for the string in a given range (r) at the column specified (cs)
    Dim rngFound As Range
    Set rngFound = rngToSearch.Find(What:=clientCode, LookIn:=xlValues, LookAt:=xlWhole)

    Fn_ClientEstValide = Not rngFound Is Nothing

    'Clean-up - 2024-08-14 @ 10:15
    Set rngFound = Nothing
    Set rngToSearch = Nothing
    
End Function

Function Fn_ValiderCourriel(ByVal adresses As String) As Boolean '2024-10-26 @ 14:30
    
    'Supporte de 0 à 2 courriels (séparés par '; ')
    
    Fn_ValiderCourriel = False
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    'Initialisation de l'expression régulière pour valider une adresse courriel
    With regex
        .Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
        .IgnoreCase = True
        .Global = False
    End With
    
    'Diviser le paremètre (courriel) en adresses individuelles
    Dim arrAdresse() As String
    arrAdresse = Split(adresses, "; ")
    
    'Vérifier chaque adresse
    Dim adresse As Variant
    For Each adresse In arrAdresse
        adresse = Trim$(adresse)
        'Passer si l'adresse est vide (Aucune adresse est aussi permis)
        If adresse <> vbNullString Then
            'Si l'adresse ne correspond pas au pattern, renvoyer Faux
            If Not regex.test(adresse) Then
                Fn_ValiderCourriel = False
                Exit Function
            End If
        End If
    Next adresse
    
    'Toutes les adresses sont valides
    Fn_ValiderCourriel = True
    
    'Nettoyer la mémoire
    Set adresse = Nothing
    Set regex = Nothing
    
End Function

Function Fn_JourMoisSpecifiqueEstIlValide(d As Long, m As Long, Y As Long) As Boolean
    'Returns TRUE or FALSE if d, m and y combined are VALID values
    
    Fn_JourMoisSpecifiqueEstIlValide = False
    
    Dim isLeapYear As Boolean
    isLeapYear = Y Mod 4 = 0 And (Y Mod 100 <> 0 Or Y Mod 400 = 0)
    
    'Last day of each month (0 to 11)
    Dim mdpm As Variant
    mdpm = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    If isLeapYear Then mdpm(1) = 29 'Adjust February for Leap Year
    
    If m < 1 Or m > 12 Or _
       d > mdpm(m - 1) Or _
       Abs(year(Date) - Y) > 75 Then
            Exit Function
    Else
        Fn_JourMoisSpecifiqueEstIlValide = True
    End If

End Function

Function Fn_AccesServeur(serverPath As String) As Boolean '2024-09-24 @ 17:14

    DoEvents
    
    Fn_AccesServeur = False
    
    'Créer un FileSystemObject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Vérifier si le fichier existe
    Dim folderExists As Boolean
    folderExists = fso.folderExists(serverPath)
    
    If folderExists Then
        Fn_AccesServeur = True
    End If
    
    'Libérer la mémoire
    Set fso = Nothing
    
End Function

Function Fn_CompleteLaDate(dateInput As String, joursArriere As Integer, joursFutur As Integer) As Variant
    
    Dim dayPart As Long
    Dim monthPart As Long
    Dim yearPart As Long
    Dim parsedDate As Date
    
    'Catch all errors
    On Error GoTo Invalid_Date
    
    'Get the current date components
    Dim defaultDay As Long
    defaultDay = day(Date)
    Dim defaultMonth As Long
    defaultMonth = month(Date)
    Dim defaultYear As Long
    defaultYear = year(Date)
    
    ' Split the input date into parts, considering different delimiters
    dateInput = Replace(Replace(Replace(dateInput, "/", "-"), ".", "-"), " ", vbNullString)
    Dim parts() As String
    parts = Split(Replace(dateInput, "-01-1900", vbNullString), "-")
    
    Select Case UBound(parts)
        Case -1
            'Nothing provided
            dayPart = defaultDay       'Use current day
            monthPart = defaultMonth   'Use current month
            yearPart = defaultYear     'Use current year
        Case 0
            'Only day provided
            dayPart = CInt(parts(0))   'Use entered day
            monthPart = defaultMonth   'Use current month
            yearPart = defaultYear     'Use current year
        Case 1
            'Day and month provided
            dayPart = CInt(parts(0))   'Use entered day
            monthPart = CInt(parts(1)) 'Use entered month
            yearPart = defaultYear     'Use current year
        Case 2
            'Day, month, and year provided
            dayPart = day(dateInput)   'Use entered day
            monthPart = month(dateInput) 'Use entered month
            yearPart = year(dateInput) 'Use entered year
            If yearPart < 100 Then
                yearPart = yearPart + 2000
            End If
        Case Else
            GoTo Invalid_Date
    End Select
    
    'Fine validation taking into consideration leap year AND 75 years (past or future)
    If Fn_JourMoisSpecifiqueEstIlValide(dayPart, monthPart, yearPart) = False Then
        GoTo Invalid_Date
    End If
    
    'Construct the full date
    parsedDate = DateSerial(yearPart, monthPart, dayPart)
    Dim joursEcart As Integer
    joursEcart = parsedDate - Date
    If joursEcart < 0 And Abs(joursEcart) > joursArriere Then
        MsgBox "Cette date NE RESPECTE PAS les paramètres de date établis" & vbNewLine & vbNewLine & _
                    "La date minimale est '" & Format$(Date - joursArriere, wsdADMIN.Range("USER_DATE_FORMAT").Value) & "'", _
                    vbCritical, "La date saisie est hors-norme - (Du " & _
                        Format$(Date - joursArriere, wsdADMIN.Range("USER_DATE_FORMAT").Value) & " au " & Format$(Date + joursFutur, wsdADMIN.Range("USER_DATE_FORMAT").Value) & ")"
        GoTo Invalid_Date
    End If
    If joursEcart > 0 And joursEcart > joursFutur Then
        MsgBox "Cette date NE RESPECTE PAS les paramètres de date établis" & vbNewLine & vbNewLine & _
                    "La date maximale est '" & Format$(Date + joursFutur, wsdADMIN.Range("USER_DATE_FORMAT").Value) & "'", _
                    vbCritical, "La date saisie est hors-norme - (Du " & _
                    Format$(Date - joursArriere, wsdADMIN.Range("USER_DATE_FORMAT").Value) & " au " & Format$(Date + joursFutur, wsdADMIN.Range("USER_DATE_FORMAT").Value) & ")"
        GoTo Invalid_Date
    End If
   
    'Return a VALID date
    Fn_CompleteLaDate = parsedDate
    
    Exit Function

Invalid_Date:
    Fn_CompleteLaDate = "Invalid Date"
    
End Function

Function Fn_TriDictionnaireParCles(dict As Object, Optional descending As Boolean = False) As Variant '2024-10-02 @ 12:02
    
    'Sort a dictionary by its keys and return keys in an array
    Dim keys() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    
    If dict Is Nothing Then
        Fn_TriDictionnaireParCles = vbNullString
        Exit Function
    End If
    
    ReDim keys(0 To dict.count - 1)
    
    Dim key As Variant
    i = 0
    For Each key In dict.keys
        keys(i) = key
        i = i + 1
    Next key
    
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If (keys(i) < keys(j) And descending) Or (keys(i) > keys(j) And Not descending) Then
                'Swap keys accordingly
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    Fn_TriDictionnaireParCles = keys
    
    'Libérer la mémoire
    Set key = Nothing
    
End Function

Function Fn_TriDictionnaireParValeurs(dict As Object, Optional descending As Boolean = False) As Variant '2024-07-11 @ 15:16
    
    'Sort a dictionary by its values and return keys in an array
    Dim keys() As Variant
    Dim values() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    
    If dict.count = 0 Then
        Exit Function
    End If
    
    ReDim keys(0 To dict.count - 1)
    ReDim values(0 To dict.count - 1)
    
    Dim key As Variant
    i = 0
    For Each key In dict.keys
        keys(i) = key
        values(i) = dict(key)
        i = i + 1
    Next key
    
    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If (values(i) < values(j) And descending) Or (values(i) > values(j) And Not descending) Then
                'Swap values
                temp = values(i)
                values(i) = values(j)
                values(j) = temp
                
                'Swap keys accordingly
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    Fn_TriDictionnaireParValeurs = keys
    
    'Libérer la mémoire
    Set key = Nothing
    
End Function

Public Function Fn_Strip_Contact_From_Client_Name(cn As String) '2024-08-15 @ 07:44

    Fn_Strip_Contact_From_Client_Name = cn
    
    'Find position of square brackets
    Dim posOSB As Integer, posCSB As Integer
    posOSB = InStr(cn, "[")
    posCSB = InStr(cn, "]")
    
    'Is there a valid structure ?
    If posOSB = 0 Or posCSB = 0 Or posCSB < posOSB Then
        Exit Function
    End If
    
    Dim cn_purged As String
    
    If posOSB > 1 Then
        Fn_Strip_Contact_From_Client_Name = Trim$(Left$(cn, posOSB - 1) & Mid$(cn, posCSB + 1))
    Else
        Fn_Strip_Contact_From_Client_Name = Trim$(Mid$(cn, posCSB + 1))
    End If
    
    'Enlever les doubles espaces
    Do While InStr(Fn_Strip_Contact_From_Client_Name, "  ")
        Fn_Strip_Contact_From_Client_Name = Replace(Fn_Strip_Contact_From_Client_Name, "  ", " ")
    Loop
    
End Function

Public Function Fn_TEC_Is_Data_Valid() As Boolean

    Fn_TEC_Is_Data_Valid = False
    
    'Validations first (one field at a time)
    
    'Professionnel ?
    If ufSaisieHeures.cmbProfessionnel.Value = vbNullString Then
        MsgBox Prompt:="Le professionnel est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        ufSaisieHeures.cmbProfessionnel.SetFocus
        Exit Function
    End If

    'Date de la charge ?
    If ufSaisieHeures.txtDate.Value = vbNullString Or IsDate(ufSaisieHeures.txtDate.Value) = False Then
        MsgBox Prompt:="La date est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        ufSaisieHeures.txtDate.SetFocus
        Exit Function
    End If

    'Nom du client & code de client ?
    If ufSaisieHeures.txtClient.Value = vbNullString Or ufSaisieHeures.txtClientID = vbNullString Then
        MsgBox Prompt:="Le client et son code sont OBLIGATOIRES !" & vbNewLine & vbNewLine & _
                       "Code de client = '" & ufSaisieHeures.txtClientID & "'" & vbNewLine & vbNewLine & _
                       "Nom du client = '" & ufSaisieHeures.txtClient.Value & "'" & vbNewLine & vbNewLine & _
                       "VOUS DEVEZ SAISIR À NOUVEAU LE CLIENT", _
               Title:="Vérifications essentielles des données du client", _
               Buttons:=vbCritical
        ufSaisieHeures.txtClient.Value = vbNullString
        ufSaisieHeures.txtClientID.Value = vbNullString
        ufSaisieHeures.txtClient.SetFocus
        Exit Function
    End If
    
    'Heures valides ?
    If ufSaisieHeures.txtHeures.Value = vbNullString Or IsNumeric(ufSaisieHeures.txtHeures.Value) = False Then
        MsgBox Prompt:="Le nombre d'heures est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        ufSaisieHeures.txtHeures.SetFocus
        Exit Function
    End If

    Fn_TEC_Is_Data_Valid = True

End Function

Public Function Fn_Get_Hourly_Rate(profID As Long, dte As Date)

    'Use the Dynamic Named Range
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names("dnrTauxHoraire").RefersToRange
    On Error GoTo 0

    'Check if the range is set correctly
    If Not rng Is Nothing Then
        Dim rowRange As Range
        Dim i As Long
        'Loop through each row in the range
        For i = rng.Rows.count To 1 Step -1
            'Set the row range
            Set rowRange = rng.Rows(i)
            If rowRange.Cells(1, 1).Value = profID Then
                If CDate(dte) >= CDate(rowRange.Cells(1, 2).Value) Then
                    Fn_Get_Hourly_Rate = rowRange.Cells(1, 3).Value
                    Exit Function
                End If
            End If
            'Loop through each cell in the row
        Next i
    Else
        MsgBox "La plage nommée 'dnrTauxHoraire' n'a pas été trouvée!", vbExclamation
    End If

    'Libérer la mémoire
    Set rng = Nothing
    Set rowRange = Nothing
    
End Function

Function Fn_TypeFacture(invNo As String) As String '2024-08-17 @ 06:55

    'Return the Type of invoice - 'C' for confirmed, 'AC' to be confirmed
    
    Dim lastUsedRow As Long
    lastUsedRow = wsdFAC_Entete.Cells(wsdFAC_Entete.Rows.count, 1).End(xlUp).Row
    Dim rngToSearch As Range
    Set rngToSearch = wsdFAC_Entete.Range("A1:A" & lastUsedRow)
    
    'Find the invNo into rngToSearch
    Dim rngFound As Range
    Set rngFound = rngToSearch.Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        Fn_TypeFacture = rngFound.offset(0, 2).Value
    Else
        Fn_TypeFacture = "C"
    End If

    'Clean-up - 2024-08-17 @ 06:55
    Set rngFound = Nothing
    Set rngToSearch = Nothing
    
End Function

Public Function Fn_Get_Tax_Rate(d As Date, taxType As String) As Double

    Dim row As Long
    Dim rate As Double
    With wsdADMIN
        For row = 18 To 11 Step -1
            If .Range("L" & row).Value = taxType Then
                If d >= .Range("M" & row).Value Then
                    rate = .Range("N" & row).Value
                    Exit For
                End If
            End If
        Next row
    End With
    
    Fn_Get_Tax_Rate = rate
    
End Function

Public Function Fn_Is_Client_Facturable(ByVal clientID As String) As Boolean '2025-07-03 @ 07:41

    Fn_Is_Client_Facturable = Len(clientID) > 2
        
End Function

Function Fn_DateEstElleValide(d As String) As Boolean

    Fn_DateEstElleValide = False
    If d = vbNullString Or IsDate(d) = False Then
        MsgBox "Une date d'écriture est obligatoire." & vbNewLine & vbNewLine & _
            "Veuillez saisir une date valide!", vbCritical, "Date Invalide"
    Else
        Fn_DateEstElleValide = True
    End If

End Function

Public Function Fn_UtilisateurWindows() As String '2025-10-19 @ 09:54

    If Len(gUtilisateurWindows) = 0 Then
        Dim buffer As String * 255
        Dim size As Long: size = 255

        If GetUserName(buffer, size) Then
            gUtilisateurWindows = Trim(Left$(buffer, size - 1))
        Else
            MsgBox "Incapable de déterminer l'utilisateur Windows !", _
                vbCritical, _
                "Fn_UtilisateurWindows"
            gUtilisateurWindows = "Unknown"
        End If
        
        If Len(gUtilisateurWindows) = 0 Then '2ème tentative
        MsgBox "Nom d'utilisateur Windows vide après tentative de récupération.", _
            vbCritical, _
            "Fn_UtilisateurWindows"
        End If
    End If

    Fn_UtilisateurWindows = gUtilisateurWindows
    
End Function

'Function Fn_UtilisateurWindows() As String '2025-06-01 @ 05:36
'
'    Dim buffer As String * 255
'    Dim size As Long: size = 255
'
'    '@Ignore UnassignedVariableUsage
'    If GetUserName(buffer, size) Then
'        '@Ignore UnassignedVariableUsage
'        gUtilisateurWindows = Trim(Left$(buffer, size - 1))
'    Else
'        MsgBox "Incapable de déterminer l'utilisateur Windows !", _
'            vbCritical, _
'            "Fn_UtilisateurWindows"
'        gUtilisateurWindows = "Unknown"
'    End If
'
'    Fn_UtilisateurWindows = gUtilisateurWindows
'
'End Function
'
'Function Fn_UtilisateurWindows() As String '2025-06-01 @ 05:48
'
'    If Len(gUtilisateurWindows) = 0 Then
'        gUtilisateurWindows = modFunctions.Fn_UtilisateurWindows()
'    End If
'
'    If Len(gUtilisateurWindows) = 0 Then
'        MsgBox "Impossible de détecter l'utilisateur Windows.", _
'            vbCritical, _
'            "modFunctions.Fn_UtilisateurWindows"
'        End
'    End If
'
'    Fn_UtilisateurWindows = gUtilisateurWindows
'
'End Function
'
Function Fn_FactureConfirmee(invNo As String) As Boolean

    Fn_FactureConfirmee = False
    
    Dim ws As Worksheet: Set ws = wsdFAC_Entete

    'Utilisation de FIND pour trouver la cellule contenant la valeur recherchée dans la colonne A
    Dim foundCell As Range
    Set foundCell = ws.Range("A:A").Find(What:=CStr(invNo), LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        If foundCell.offset(0, 2).Value = "C" Then
            Fn_FactureConfirmee = True
        End If
    Else
        Fn_FactureConfirmee = False
    End If

    'Libérer la mémoire
    Set foundCell = Nothing
    Set ws = Nothing

End Function

Function Fn_SaisieEJBalance() As Boolean

    Fn_SaisieEJBalance = False
    If wshGL_EJ.Range("H26").Value <> wshGL_EJ.Range("I26").Value Then
        MsgBox "Votre écriture ne balance pas." & vbNewLine & vbNewLine & _
            "Débits = " & wshGL_EJ.Range("H26").Value & " et Crédits = " & wshGL_EJ.Range("I26").Value & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée.", vbCritical, "Veuillez vérifier votre écriture!"
    Else
        Fn_SaisieEJBalance = True
    End If

End Function

Function Fn_SaisieDEBBalance() As Boolean

    Fn_SaisieDEBBalance = False
    If CCur(wshDEB_Saisie.Range("O6").Value) <> CCur(wshDEB_Saisie.Range("I26").Value) Then
        MsgBox "Votre transaction ne balance pas." & vbNewLine & vbNewLine & _
            "Total saisi = " & Format$(wshDEB_Saisie.Range("O6").Value, "#,##0.00 $") _
            & " vs. Ventilation = " & Format$(wshDEB_Saisie.Range("I26").Value, "#,##0.00 $") _
            & vbNewLine & vbNewLine & "Elle n'est donc pas reportée.", _
            vbCritical, "Veuillez vérifier votre écriture!"
    Else
        Fn_SaisieDEBBalance = True
    End If

End Function

Function Fn_SaisieEJEstValide(rmax As Long) As Boolean

    Fn_SaisieEJEstValide = True 'Optimist
    If rmax <= 9 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        Fn_SaisieEJEstValide = False
    End If

End Function

Function Fn_SaisieDEBEstElleValide(rmax As Long) As Boolean

    Fn_SaisieDEBEstElleValide = True 'Optimist
    If rmax < 9 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        Fn_SaisieDEBEstElleValide = False
    End If
    
    Dim i As Long
    For i = 9 To rmax
        If wshDEB_Saisie.Range("E" & i).Value <> vbNullString Then
            If wshDEB_Saisie.Range("N" & i).Value = vbNullString Then
                MsgBox _
                    Prompt:="Il existe une ligne avec un compte, sans montant !", _
                    Title:="L'entrée ne peut être acceptée dans son état actuel", _
                    Buttons:=vbInformation
                Fn_SaisieDEBEstElleValide = False
            End If
        End If
    Next i

End Function

Public Function Fn_ChaineRemplie(s As String, fillCaracter As String, length As Long, leftOrRight As String) As String

    Dim charactersNeeded As Long
    charactersNeeded = length - Len(s)
    
    Dim paddedString As String
    If charactersNeeded > 0 Then
        If leftOrRight = "R" Then
            paddedString = s & String(charactersNeeded, fillCaracter)
        Else
            paddedString = String(charactersNeeded, fillCaracter) & s
        End If
    Else
        paddedString = s
    End If

    Fn_ChaineRemplie = paddedString
        
End Function

Function Fn_ProchainNumeroFacture() As String '2024-09-17 @ 14:00

    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim strLastInvoice As String
    strLastInvoice = ws.Cells(lastUsedRow, 1).Value
    If strLastInvoice <> vbNullString Then
        strLastInvoice = Right$(strLastInvoice, Len(strLastInvoice) - 3)
    Else
        MsgBox "Problème avec les dernières lignes de la" & _
                vbNewLine & vbNewLine & "feuille 'wsdFAC_Entete'" & _
                vbNewLine & vbNewLine & "Veuillez contacter le développeur", _
                vbOKOnly, "Structure invalide dans 'wsdFAC_Entete'"
    End If
    Fn_ProchainNumeroFacture = strLastInvoice + 1

    'Libérer la mémoire
    Set ws = Nothing
    
End Function

Function Fn_PlanComptableTableau2D(nbCol As Long) As Variant '2024-06-07 @ 07:31

    Debug.Assert nbCol >= 1 And nbCol <= 4 '2024-07-31 @ 19:26
    
    'Reference the named range
    Dim planComptable As Range: Set planComptable = wsdADMIN.Range("dnrPlanComptable_All")
    
    'Iterate through each row of the named range
    Dim rowNum As Long, row As Range, rowRange As Range
    Dim arr() As String
    If nbCol = 1 Then
        ReDim arr(1 To planComptable.Rows.count) As String '1D array
    Else
        ReDim arr(1 To planComptable.Rows.count, 1 To nbCol) As String '2D array
    End If
    For rowNum = 1 To planComptable.Rows.count
        'Get the entire row as a range
        Set rowRange = planComptable.Rows(rowNum)
        'Process each cell in the row
        For Each row In rowRange.Rows
            If nbCol = 1 Then
                arr(rowNum) = row.Cells(1, 2)
            ElseIf nbCol = 2 Then
                arr(rowNum, 1) = row.Cells(1, 2)
                arr(rowNum, 2) = row.Cells(1, 1)
            ElseIf nbCol = 3 Then
                arr(rowNum, 1) = row.Cells(1, 2)
                arr(rowNum, 2) = row.Cells(1, 1)
                arr(rowNum, 3) = row.Cells(1, 3)
            Else
                arr(rowNum, 1) = row.Cells(1, 2)
                arr(rowNum, 2) = row.Cells(1, 1)
                arr(rowNum, 3) = row.Cells(1, 3)
                arr(rowNum, 4) = row.Cells(1, 4)
            End If
        Next row
    Next rowNum
    
    Fn_PlanComptableTableau2D = arr
    
    'Libérer la mémoire
    Set planComptable = Nothing
    Set row = Nothing
    Set rowRange = Nothing
    
End Function

Function Fn_NomClientFeuilleBDCients(cc As String) As String

    Dim ws As Worksheet
    Dim foundCell As Range
    
    Set ws = wsdBD_Clients
    
    'Recherche le code de client dans la colonne B
    Set foundCell = ws.Columns("B").Find(What:=cc, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        'Si trouvé, retourner le nom du client correspondant, 1 colonne à gauche
        Fn_NomClientFeuilleBDCients = foundCell.offset(0, -1).Value
    Else
        Fn_NomClientFeuilleBDCients = "Client non trouvé (invalide)"
    End If
    
    'Libérer la mémoire
    Set foundCell = Nothing
    Set ws = Nothing
    
End Function

Function Fn_LigneClientAPartirDuClientID(codeClient As String, ws As Worksheet) As Variant

    'Recherche de l'ID du client dans la colonne B
    Dim rangeID As Range:
    Set rangeID = ws.Columns("B") 'Contient les ID des clients
    
    'Utilisation de Find pour localiser l'ID client
    Dim foundCells As Range
    Set foundCells = rangeID.Find(What:=codeClient, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Si l'ID client est trouvé
    Dim ligneTrouvee As Long
    If Not foundCells Is Nothing Then
        'Obtenir la ligne où se trouve l'ID client
        ligneTrouvee = foundCells.row
        
        'Extraire toutes les données (colonnes) de la ligne trouvée
        Dim clientData As Variant
        clientData = ws.Rows(ligneTrouvee).Value
        
        'Retourner les données du client (ligne entière)
        Fn_LigneClientAPartirDuClientID = clientData
    Else
        'Si le client n'est pas trouvé, retourner une valeur vide ou une erreur
        Fn_LigneClientAPartirDuClientID = CVErr(xlErrNA) 'Retourne #N/A pour indiquer que le client n'est pas trouvé
    End If
    
    'Libérer la mémoire
    Set foundCells = Nothing
    Set rangeID = Nothing
    
End Function

Function Fn_ChaineSansAccents(ByVal text As String) As String

    'Liste des caractères accentués et leurs équivalents sans accents
    Dim AccChars As String
    AccChars = "ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
    Dim RegChars As String
    RegChars = "AAAAAACEEEEIIIINOOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"

    'Remplacer les accents par des caractères non accentués
    Dim i As Long
    For i = 1 To Len(AccChars)
        text = Replace(text, Mid$(AccChars, i, 1), Mid$(RegChars, i, 1))
    Next i

    Fn_ChaineSansAccents = text
    
End Function

Public Function Fn_Get_Current_Region(ByVal dataRange As Range, Optional HeaderSize As Long = 1) As Range

    Set Fn_Get_Current_Region = dataRange.CurrentRegion
    If HeaderSize > 0 Then
        With Fn_Get_Current_Region
            'Remove the header
            Set Fn_Get_Current_Region = .Offset(HeaderSize).Resize(.Rows.count - HeaderSize)
            Debug.Print "#060 - " & Fn_Get_Current_Region.Address
        End With
    End If
    
    'Libérer la mémoire
    Set Fn_Get_Current_Region = Nothing
    
End Function

Public Function Fn_Convert_Value_Boolean_To_Text(val As Boolean) As String

    Select Case val
        Case 0, "False", "Faux", "FAUX" 'False
            Fn_Convert_Value_Boolean_To_Text = "FAUX"
        Case -1, "True", "Vrai", "VRAI" 'True"
            Fn_Convert_Value_Boolean_To_Text = "VRAI"
        Case Else
            MsgBox val & " est une valeur INVALIDE !"
    End Select

End Function

Function Fn_ChaineValideAvecPlage(searchString As String, rng As Range) As Boolean

    On Error Resume Next
    Fn_ChaineValideAvecPlage = Not IsError(Application.Match(searchString, rng, 0))
    On Error GoTo 0
    
End Function

'Fonction de tri rapide (QuickSort) pour trier un tableau
Sub TrierBubble(arr As Variant, ByVal first As Long, ByVal last As Long) '2024-09-05 @ 05:09
    
    Dim pivot As Variant, tmp As Variant
    Dim i As Long, j As Long
    
    If first < last Then
        pivot = arr((first + last) \ 2)
        i = first
        j = last
        Do
            Do While arr(i) < pivot: i = i + 1: Loop
            Do While arr(j) > pivot: j = j - 1: Loop
            If i <= j Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
                i = i + 1
                j = j - 1
            End If
        Loop While i <= j
        If first < j Then Call TrierBubble(arr, first, j)
        If i < last Then Call TrierBubble(arr, i, last)
    End If
    
End Sub

Function Fn_FractionHeureEstValide(valeur As Currency) As Boolean

    'Tableau des valeurs permises : dixièmes d'heures et quarts d'heure
    Dim valeursPermises As Variant
    valeursPermises = Array(0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 0.25, 0.75)
    
    Dim i As Integer
    Fn_FractionHeureEstValide = False 'Initialisation à Faux
    
    'Parcourir les valeurs permises
    Dim fraction As Double
    fraction = valeur - Int(valeur)
    
    For i = LBound(valeursPermises) To UBound(valeursPermises)
        If Round(fraction, 2) = valeursPermises(i) Then
            Fn_FractionHeureEstValide = True 'La fraction est valide
            Exit Function
        End If
    Next i
    
End Function

Function Fn_DatePremierJourTrimestrePrecedent(d As Date) As Date

    'Cette fonction calcule le premier jour du trimestre pour une date de fin de trimestre (TPS/TVQ)
    Dim dateTroisMoisAvant As Date
    
    'Reculer de trois mois à partir de la date saisie
    dateTroisMoisAvant = DateAdd("m", -2, d)
    
    'Fixer le jour au PREMIER du mois obtenu
    Fn_DatePremierJourTrimestrePrecedent = DateSerial(year(dateTroisMoisAvant), month(dateTroisMoisAvant), 1)
    
End Function

Function Fn_DateDuLundi(d As Date)

    Fn_DateDuLundi = d - (Weekday(d, vbMonday) - 1)

End Function

Function Fn_ChaineNettoyeeCaracteresSpeciaux(s As String) '2024-11-07 @ 16:57

    Fn_ChaineNettoyeeCaracteresSpeciaux = s
    
    'Supprimer les retours à la ligne, les sauts de ligne et les espaces inutiles
    Fn_ChaineNettoyeeCaracteresSpeciaux = Trim$(Replace(Replace(Replace(s, vbCrLf, vbNullString), vbCr, vbNullString), vbLf, vbNullString))

End Function

'Fonction pour vérifier si un fichier ou un dossier existe
Private Function Fn_Chemin_Existe(ByVal chemin As String) As Boolean

    On Error Resume Next
    Fn_Chemin_Existe = (Dir(chemin) <> vbNullString)
    On Error GoTo 0
    
End Function

Function Fn_NoCompteAPartirIndicateurCompte(ByVal indic As Variant) As String

    'Plage où sont situés les liens (indicateur/no de GL)
    Dim plage As Range
    Set plage = wsdADMIN.Range("D44:F63")
    
    'Parcourir chaque cellule dans la première colonne de la plage
    Dim cellule As Range
    For Each cellule In plage.Columns(1).Cells
        If cellule.Value = indic Then
            'Retourner la valeur de la deuxième colonne de la ligne trouvée
            Fn_NoCompteAPartirIndicateurCompte = cellule.offset(0, 1).Value
            Exit Function
        End If
    Next cellule
    
    'Si la valeur n'est pas trouvée
    Fn_NoCompteAPartirIndicateurCompte = "Non trouvé"

End Function

Function Fn_ObtenirPaiementsPourUneFacture(invNo As String, dateLimite As Date) As Currency

    Dim ws As Worksheet: Set ws = wsdENC_Details
    
    Dim premièreCellule As Range, celluleTrouvée As Range
    Dim ligne As Long
    
    'Rechercher la première cellule avec le numéro de facture
    Set premièreCellule = ws.Columns(2).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    Fn_ObtenirPaiementsPourUneFacture = 0
    
    If Not premièreCellule Is Nothing Then
        'Démarrer la recherche à partir de la première cellule trouvée
        Set celluleTrouvée = premièreCellule
        Do
            'Obtenir la ligne correspondante
            ligne = celluleTrouvée.row
            'Vérifier la date dans la colonne appropriée
            If IsDate(ws.Cells(ligne, 4).Value) And ws.Cells(ligne, 4).Value <= dateLimite Then
                'Additionner le montant du paiement
                Fn_ObtenirPaiementsPourUneFacture = Fn_ObtenirPaiementsPourUneFacture + ws.Cells(ligne, 5).Value
            End If
            'Rechercher la prochaine occurrence
            Set celluleTrouvée = ws.Columns(2).FindNext(celluleTrouvée)
        Loop While Not celluleTrouvée Is Nothing And celluleTrouvée.Address <> premièreCellule.Address
    End If

End Function

Function Fn_ObtenirRegularisationsFacture(invNo As String, dateLimite As Date) As Currency

    Dim ws As Worksheet: Set ws = wsdCC_Regularisations
    
    Dim premièreCellule As Range, celluleTrouvée As Range
    Dim ligne As Long
    
    'Rechercher la première cellule avec le numéro de facture
    Set premièreCellule = ws.Columns(fREGULInvNo).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    Fn_ObtenirRegularisationsFacture = 0
    
    If Not premièreCellule Is Nothing Then
        'Démarrer la recherche à partir de la première cellule trouvée
        Set celluleTrouvée = premièreCellule
        Do
            'Obtenir la ligne correspondante
            ligne = celluleTrouvée.row
            'Vérifier la date dans la colonne appropriée
            If IsDate(ws.Cells(ligne, fREGULDate).Value) And ws.Cells(ligne, fREGULDate).Value <= dateLimite Then
                'Additionner les cellules pertinentes
                Fn_ObtenirRegularisationsFacture = Fn_ObtenirRegularisationsFacture + _
                                                        ws.Cells(ligne, fREGULHono).Value + _
                                                        ws.Cells(ligne, fREGULFrais).Value + _
                                                        ws.Cells(ligne, fREGULTPS).Value + _
                                                        ws.Cells(ligne, fREGULTVQ).Value
            End If
            'Rechercher la prochaine occurrence
            Set celluleTrouvée = ws.Columns(fREGULInvNo).FindNext(celluleTrouvée)
        Loop While Not celluleTrouvée Is Nothing And celluleTrouvée.Address <> premièreCellule.Address
    End If

End Function

Function Fn_ExtraireSecondesChaineLog(chaine As String) As Double
    
    Dim pos As Integer
    Dim secondes As String
    
    chaine = Replace(chaine, ".", ",")
    chaine = Replace(chaine, "'", vbNullString)
    
    If InStr(chaine, " = ") > 0 Then
        chaine = Right$(chaine, Len(chaine) - InStr(chaine, " = ") - 2)
    End If
    
    'Trouve la position de " secondes"
    pos = InStr(chaine, " secondes")
    If pos > 0 Then
        secondes = Left$(chaine, pos - 1)
        Fn_ExtraireSecondesChaineLog = CStr(secondes)
    Else
        'Si " secondes" n'est pas trouvé, retourne une chaîne vide
        Fn_ExtraireSecondesChaineLog = "0"
    End If
    
End Function

'Fonction pour centraliser tous les messages de l'application - 2024-12-29 @ 07:37
Function Fn_AppMsgBox(message As String _
                 , Optional boutons As VbMsgBoxStyle = vbOKOnly _
                 , Optional titre As String = vbNullString) As VbMsgBoxResult
                 
    Fn_AppMsgBox = MsgBox(message, boutons, titre)
                 
End Function

Sub zz_TesterAppMsgBox() '@TODO

    Dim r As VbMsgBoxResult
    r = Fn_AppMsgBox("Voulez-vous continuer ?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirmation avant de continuer")

    Debug.Print "#090 - " & r
    
End Sub

Function Fn_ExtraireNomFichier(path As String) As String

    Dim parts() As String
    parts = Split(path, "\")
    
    Fn_ExtraireNomFichier = parts(UBound(parts, 1))

End Function

Sub ConvertirEnNumerique(rng As Range)

    Dim cell As Range
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            cell.Value = CCur(cell.Value)
        Else
            cell.Value = CCur(Replace(cell.Value, " ", vbNullString))
        End If
    Next cell
    
End Sub

Function Fn_RangeeFactureSpecifique(ws As Worksheet, numFacture As String) As Long

    Dim rng As Range
    Set rng = ws.Range("A:A").Find(What:=numFacture, LookAt:=xlWhole)

    If Not rng Is Nothing Then
        Fn_RangeeFactureSpecifique = rng.row
    Else
        Fn_RangeeFactureSpecifique = -1 ' Retourne -1 si la facture n'est pas trouvée
    End If
    
End Function

Function Fn_RangeeAPartirNumeroColonne1(ws As Worksheet, noEntrée As Long) As Variant

    'Définition de la feuille contenant les données
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    Dim plageRecherche As Range
    Set plageRecherche = ws.Range("A2:A" & lastUsedRow)

    'Recherche des lignes correspondantes
    Dim plageResultat As Range
    Dim cell As Range
    For Each cell In plageRecherche
        If cell.Value = noEntrée Then
            If plageResultat Is Nothing Then
                Set plageResultat = cell.EntireRow
            Else
                Set plageResultat = Union(plageResultat, cell.EntireRow)
            End If
        End If
    Next cell

    ' Vérifier si des résultats ont été trouvés
    If plageResultat Is Nothing Then
        Fn_RangeeAPartirNumeroColonne1 = False ' Aucun résultat trouvé
        Exit Function
    End If

    'Convertir la plage trouvée en tableau VBA
    Dim nbLignes As Integer, nbColonnes As Integer
    Dim data As Variant
    nbLignes = plageResultat.Rows.count
    nbColonnes = plageResultat.Columns.count
    data = plageResultat.Value ' Stocker la plage dans un tableau Variant

    ' Retourner le tableau des résultats
    Fn_RangeeAPartirNumeroColonne1 = data
    
End Function

Function Fn_EstLigneSelectionnee(ByVal lb As Object) As Boolean

    Dim i As Integer
    Fn_EstLigneSelectionnee = False 'Par défaut, aucune ligne n'est sélectionnée

    'Vérifier toutes les lignes du ListBox
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            Fn_EstLigneSelectionnee = True 'Si une ligne est sélectionnée
            Exit Function 'Quitter dès qu'une sélection est trouvée
        End If
    Next i
    
End Function

Function Fn_ExtraireTransactionsPourUnCompte(rngResultAll As Range, compte As String) As Variant

    Dim rFiltre As Range
    Dim i As Long, j As Long
    Dim count As Long
    
    'Déduire la feuille de travail à partir de rngResultAll
    Dim ws As Worksheet
    Set ws = rngResultAll.Parent
    
    'Définir une cellule pour le critère de filtrage temporaire
    Dim Critere As Range
    Set Critere = ws.Range("AN2:AN3") 'A/F # 3 dans Gl_Trans
    'Nettoyer les anciennes valeurs du critère
    Critere.Cells.ClearContents
    Critere.Cells(1, 1).Value = "NoCompte" 'Titre de la colonne à filtrer (Adapter si nécessaire)
    Critere.Cells(2, 1).Value = compte
    
    'Nettoyer la plage pour recevoir la copie
    ws.Range("AP1").CurrentRegion.offset(1, 0).Clear
    
    'Appliquer AdvancedFilter pour obtenir uniquement les lignes correspondant au compte
    rngResultAll.AdvancedFilter action:=xlFilterCopy, _
                                criteriaRange:=Critere, _
                                CopyToRange:=ws.Range("AP1:AX1"), _
                                Unique:=False
    
    'Récupérer la plage filtrée
    Set rFiltre = ws.Range("AP1").CurrentRegion
    
    'Vérifier qu'il y a des lignes filtrées
    Dim LignesCorrespondantes() As Variant
    If rFiltre.Rows.count > 1 Then 'Lignes après l'en-tête
        'Initialiser le tableau pour contenir les lignes pertinentes
        ReDim LignesCorrespondantes(1 To rFiltre.Rows.count - 1, 1 To rFiltre.Columns.count)
        
        'Remplir le tableau avec les données filtrées
        count = 1
        For i = 2 To rFiltre.Rows.count ' Ignorer l'en-tête
            For j = 1 To rFiltre.Columns.count
                LignesCorrespondantes(count, j) = rFiltre.Cells(i, j).Value
            Next j
            count = count + 1
        Next i
    Else
        LignesCorrespondantes = Array()
    End If
    
    'Retourner le tableau contenant les lignes filtrées ou Vide (Array())
    Fn_ExtraireTransactionsPourUnCompte = LignesCorrespondantes
    
    'Libérer la mémoire
    Set rFiltre = Nothing
    
End Function

Function Fn_EstBissextile(annee As Integer) As Boolean '2025-03-02 @ 10:21

    ' Vérifie si l'année est bissextile
    Fn_EstBissextile = ((annee Mod 4 = 0 And annee Mod 100 <> 0) Or annee Mod 400 = 0)
    
End Function

Function Fn_ValiderDateDernierJourDuMois(Y As Integer, m As Integer, d As Integer) As String '2025-03-02 @ 11:00

    'Le mois est-il valide (max 12)
    If m > 12 Then
        Fn_ValiderDateDernierJourDuMois = vbNullString
        Exit Function
    End If
    'Quel est le dernier jour de ce mois ?
    Dim dernierJourDuMois As Integer
    dernierJourDuMois = day(DateSerial(Y, m + 1, 0))
    
    'Vérification additionnelle pour le mois de février (années bissextiles)
    Dim isLeapYear As Boolean
    isLeapYear = False
    If m = 2 Then
        isLeapYear = Fn_EstBissextile(Y)
        If isLeapYear Then
            dernierJourDuMois = 29
        End If
    End If
    
    'Valider le jour
    If d > dernierJourDuMois Then
        Dim message As String, titre As String
        message = "La combinaison jour (" & d & ") et mois (" & m & ") n'existe pas"
        titre = "Date invalide"
        If m = 2 Then
            titre = titre & IIf(isLeapYear, vbNullString, " pour l'année " & Y)
        End If
        MsgBox message, vbExclamation, titre
        Fn_ValiderDateDernierJourDuMois = vbNullString 'Si le jour n'est pas valide, on retourne une chaîne vide
        Exit Function
    Else
        Fn_ValiderDateDernierJourDuMois = DateSerial(Y, m, d)
    End If
    
End Function

Function Fn_ExclureTransaction(source As String) As Boolean '2025-03-06 @ 08:04

    'Liste des sources à exclure
    Dim Exclusions As Variant
    Exclusions = Array("DÉBOURSÉ:", "DÉPÔT DE CLIENT:", "ENCAISSEMENT:", "FACTURE:", "RÉGULARISATION:", "RENVERSEMENT:", "RENVERSÉE par ")
    
    Dim i As Integer
    For i = LBound(Exclusions) To UBound(Exclusions)
        If Left$(source, Len(Exclusions(i))) = Exclusions(i) Then
            Fn_ExclureTransaction = True
            Exit Function
        End If
    Next i
    
    Fn_ExclureTransaction = False
    
End Function

Function Fn_PremierJourTrimestreFiscal(dateMax As Date) As Date

    Dim annee As Integer
    Dim mois As Integer
    Dim MoisFinAnneeFinanciere As Integer
    Dim decalage As Integer

    mois = month(dateMax)
    MoisFinAnneeFinanciere = wsdADMIN.Range("MoisFinAnnéeFinancière")

    If mois < MoisFinAnneeFinanciere + 1 Then
        annee = year(dateMax) - 1
        decalage = Int((mois + 12 - (MoisFinAnneeFinanciere + 1)) / 3) * 3
    Else
        annee = year(dateMax)
        decalage = Int((mois - (MoisFinAnneeFinanciere + 1)) / 3) * 3
    End If

    Fn_PremierJourTrimestreFiscal = DateSerial(annee, MoisFinAnneeFinanciere + 1 + decalage, 1)
    
End Function

Function Fn_PremierJourAnneeFinanciere(maxDate As Date) As Date

    Dim dt As Date
    Dim annee As Integer

    'Calcul de l'année du début d'exercice
    If month(maxDate) > wsdADMIN.Range("MoisFinAnnéeFinancière") Then
        annee = year(maxDate)
    Else
        annee = year(maxDate) - 1
    End If

    'Retourner le 1er jour du mois suivant la fin d'exercice
    If wsdADMIN.Range("MoisFinAnnéeFinancière") = 12 Then
        dt = DateSerial(annee + 1, 1, 1)
    Else
        dt = DateSerial(annee, wsdADMIN.Range("MoisFinAnnéeFinancière") + 1, 1)
    End If

    Fn_PremierJourAnneeFinanciere = dt
    
End Function

Function Fn_DernierJourAnneeFinanciere(dateReference As Date) As Date '2025-08-14 @ 20:05

    Dim dt As Date
    Dim finExercice As Date
    Dim anneeCloture As Integer
    Dim dernierJour As Integer

    'Calcul de l'année du début d'exercice
    If month(dateReference) > wsdADMIN.Range("MoisFinAnnéeFinancière") Then
        anneeCloture = year(dateReference) + 1
    Else
        anneeCloture = year(dateReference)
    End If

    'Retourner le 1er jour du mois suivant la fin d'exercice
    dernierJour = day(DateSerial(anneeCloture, wsdADMIN.Range("MoisFinAnnéeFinancière") + 1, 0))
    finExercice = DateSerial(anneeCloture, wsdADMIN.Range("MoisFinAnnéeFinancière"), dernierJour)

    Fn_DernierJourAnneeFinanciere = finExercice
    
End Function

Function Fn_TableauContientDesDonnees(lo As ListObject) As Boolean '2025-06-01 @ 06:41

    If lo.DataBodyRange Is Nothing Then
        Fn_TableauContientDesDonnees = False
    Else
        Fn_TableauContientDesDonnees = (Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0)
    End If
    
End Function

Function Fn_DateNormalisee(chaine As String) As Variant '2025-06-12 @ 08:22

    Dim j As Integer, m As Integer, a As Integer
    Dim parties() As String
    Dim nbParties As Integer
    Dim resultDate As Date
    
    chaine = Trim(chaine)
    If chaine = vbNullString Then
        Fn_DateNormalisee = CVErr(xlErrValue)
        Exit Function
    End If

    parties = Split(Replace(chaine, "-", "/"), "/")
    nbParties = UBound(parties) - LBound(parties) + 1

    On Error GoTo erreur

    Select Case nbParties
        Case 1 ' Juste le jour
            j = CInt(parties(0))
            m = month(Date)
            a = year(Date)
        
        Case 2 ' Jour et mois
            j = CInt(parties(0))
            m = CInt(parties(1))
            a = year(Date)

        Case 3 ' Date complète
            ' Tenter d’interpréter selon plusieurs ordres
            If Len(parties(0)) = 4 Then
                ' Format aaaa/mm/jj
                a = CInt(parties(0))
                m = CInt(parties(1))
                j = CInt(parties(2))
            ElseIf Len(parties(2)) = 4 Then
                ' Format jj/mm/aaaa
                j = CInt(parties(0))
                m = CInt(parties(1))
                a = CInt(parties(2))
            Else
                ' Ambigu – considérer jj/mm/aa (par défaut)
                j = CInt(parties(0))
                m = CInt(parties(1))
                a = CInt(parties(2))
                If a < 100 Then a = 2000 + a ' corriger pour 2 chiffres
            End If
        
        Case Else
            Fn_DateNormalisee = CVErr(xlErrValue)
            Exit Function
    End Select

    'Validation via DateSerial (gère les bissextiles)
    resultDate = DateSerial(a, m, j)
    Fn_DateNormalisee = resultDate
    Exit Function

erreur:
    Fn_DateNormalisee = CVErr(xlErrValue)
    
End Function

'Function Fn_MinutesDepuisDerniereActivite() As Double '2025-07-01 @ 14:06
'
'    If gDerniereActivite = 0 Then
'        Fn_MinutesDepuisDerniereActivite = 0
'    Else
'        Fn_MinutesDepuisDerniereActivite = DateDiff("s", gDerniereActivite, Now) / 60
'    End If
'
'End Function
'
'Public Function GetProchaineFermeture() As Date '2025-07-02 @ 09:38
'
'    GetProchaineFermeture = Now + TimeSerial(0, 0, gDELAI_GRACE_SECONDES)
'
'End Function
'
Public Function EstChampModifie(champ As String, valeurOrigine As String) As Boolean '2025-07-03 @ 07:15

    EstChampModifie = (Trim(champ & vbNullString) <> Trim(valeurOrigine & vbNullString))
    
End Function

Public Function Fn_NomFeuilleActive() As String '2025-07-03 @ 10:18

    On Error Resume Next
    Fn_NomFeuilleActive = ActiveSheet.Name
    On Error GoTo 0
    
End Function

Function Fn_ConsidereOuPasCetteEcriture(ByVal source As String) As Boolean '2025-03-03 @ 10:21

    'Variable pour vérifier si la transaction est valide
    Dim aImprimer As Boolean
    aImprimer = False

    'Traitement de la transaction selon fGlTSource et l'état des cases
    If InStr(source, "DÉBOURSÉ:") = 1 Or InStr(source, "RENV/DÉBOURSÉ:") = 1 Then
        If ufGL_Rapport.chkDebourse.Value = True Then aImprimer = True
    ElseIf InStr(source, "DÉPÔT DE CLIENT:") = 1 Then
        If ufGL_Rapport.chkDepotClient.Value = True Then aImprimer = True
    ElseIf InStr(source, "ENCAISSEMENT:") = 1 Then
        If ufGL_Rapport.chkEncaissement.Value = True Then aImprimer = True
    ElseIf InStr(source, "FACTURE:") = 1 Then
        If ufGL_Rapport.chkFacture.Value = True Then aImprimer = True
    ElseIf InStr(source, "RÉGULARISATION:") = 1 Then
        If ufGL_Rapport.chkRegularisation.Value = True Then aImprimer = True
    ElseIf InStr(source, "Clôture Annuelle") = 1 Then
        If ufGL_Rapport.chkEcrCloture.Value = True Then aImprimer = True
    Else
        If ufGL_Rapport.chkEJ.Value = True Then aImprimer = True
    End If

    'Retourne True si la transaction doit être traitée, sinon False
    Fn_ConsidereOuPasCetteEcriture = aImprimer
    
End Function

Function Fn_ExtraireVersionDepuisNomClasseur() As String '2025-08-12 @ 16:04

    Dim posV As Long
    Dim posExt As Long
    Dim versionBrute As String
    Dim nomFichier As String
    nomFichier = ThisWorkbook.Name

    posV = InStr(nomFichier, "APP_v")
    If posV = 0 Then
        Fn_ExtraireVersionDepuisNomClasseur = ""
        Exit Function
    End If

    versionBrute = Mid(nomFichier, posV + 5)
    posExt = InStr(versionBrute, ".xlsb")
    If posExt > 0 Then
        versionBrute = Left(versionBrute, posExt - 1)
    End If

    Fn_ExtraireVersionDepuisNomClasseur = versionBrute
    
End Function

Function Fn_DateMoinsUnAn(dateInitiale As Date) As Date '2025-08-13 @ 16:46

    Dim anneeCible As Integer
    Dim mois As Integer
    Dim jour As Integer
    Dim tentative As Date

    anneeCible = year(dateInitiale) - 1
    mois = month(dateInitiale)
    jour = day(dateInitiale)

    'Cas spécial : dernier jour de février
    If mois = 2 And jour = day(DateSerial(year(dateInitiale), 3, 0)) Then
        'On retourne le dernier jour de février de l'année précédente
        tentative = DateSerial(anneeCible, 3, 0)
    Else
        On Error Resume Next
        tentative = DateSerial(anneeCible, mois, jour)
        If Err.Number <> 0 Then
            Err.Clear
            tentative = DateSerial(anneeCible, mois + 1, 0)
        End If
        On Error GoTo 0
    End If

    Fn_DateMoinsUnAn = tentative
    
End Function

Function Fn_ObtenirOuCreerFeuille(ByVal nomFeuille As String) As Worksheet '2025-08-14 @ 09:47

    Dim ws As Worksheet
    Dim FeuilleExiste As Boolean
    Dim Sh As Worksheet

    FeuilleExiste = False
    'Vérifie si la feuille existe
    For Each Sh In ThisWorkbook.Worksheets
        If Sh.Name = nomFeuille Then
            Set ws = Sh
            FeuilleExiste = True
            Exit For
        End If
    Next Sh

    'Si elle n'existe pas, on la crée
    If Not FeuilleExiste Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        On Error Resume Next
        ws.Name = nomFeuille
        If Err.Number <> 0 Then
            MsgBox "Impossible de nommer la feuille '" & nomFeuille & "'.", vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    End If

    'Retourne la feuille
    Set Fn_ObtenirOuCreerFeuille = ws
    
End Function

Function Fn_ContexteActifComplet() As String '2025-10-30 @ 06:20

    Dim nomFeuille As String: nomFeuille = ""
    Dim nomFormulaire As String: nomFormulaire = ""
    Dim nomControle As String: nomControle = ""

    'Feuille active (sécurisée)
    On Error Resume Next
    If Not Application.ActiveSheet Is Nothing Then
        nomFeuille = Application.ActiveSheet.Name
    End If
    On Error GoTo 0

    'Formulaire actif + contrôle actif
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.Visible Then
            nomFormulaire = uf.Name
            On Error Resume Next
            If Not uf.ActiveControl Is Nothing Then
                nomControle = uf.ActiveControl.Name
            End If
            On Error GoTo 0
            Exit For
        End If
    Next uf

    'Construction du message
    Dim message As String
    If nomFeuille <> vbNullString Then
        message = message & "Feuille: " & nomFeuille & " / "
    End If
    If nomFormulaire <> vbNullString Then
        message = message & "Formulaire: " & nomFormulaire & " / "
    End If
    If nomControle <> vbNullString Then
        message = message & "Contrôle: " & nomControle
    End If
    If Right(message, 3) = " / " Then
        message = Left(message, Len(message) - 3)
    End If
    Fn_ContexteActifComplet = message
    
End Function

