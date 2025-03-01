Attribute VB_Name = "modFunctions"
Option Explicit

'API pour code d'utilisateur
Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function Fn_GetID_From_Initials(i As String)

    Dim cell As Range
    
    For Each cell In wshAdmin.Range("dnrProf_All")
        If cell.Value2 = i Then
            Fn_GetID_From_Initials = cell.offset(0, 1).value
            Exit Function
        End If
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Function Fn_Get_Prof_From_ProfID(i As Long)

    Dim cell As Range
    
    For Each cell In wshAdmin.Range("dnrProf_All")
        If cell.Value2 = i Then
            Fn_Get_Prof_From_ProfID = cell.offset(0, -1).value
            Exit Function
        End If
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Function Fn_Get_A_Row_From_A_Worksheet(feuille As String, cle As Variant, cleCol As Integer) As Variant

    'Feuille à rechercher
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(feuille)
    
    'Charger les données en mémoire
    Dim allData As Variant
    allData = ws.usedRange.value

    'Parcourir les données pour trouver la valeur
    Dim resultArray() As Variant
    Dim i As Long
    For i = 1 To UBound(allData, 1)
        If allData(i, cleCol) = cle Then
            'Ligne est trouvée alors on copie toutes les colonnes dans le tableau résultat
            resultArray = Application.index(allData, i, 0)
            Fn_Get_A_Row_From_A_Worksheet = resultArray
            Exit Function
        End If
    Next i
    
    'Si aucune correspondance n'a été trouvée, retourner une valeur vide
    Fn_Get_A_Row_From_A_Worksheet = CVErr(xlErrValue)
    
End Function

Function Fn_GetID_From_Client_Name(nomClient As String) '2024-02-14 @ 06:07

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFunctions:Fn_GetID_From_Client_Name", nomClient, 0)
    
    Dim ws As Worksheet: Set ws = wshBD_Clients
    
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
        Fn_GetID_From_Client_Name = result
        ufSaisieHeures.txtClientID.value = result
    Else
        MsgBox "Impossible de retrouver le nom du client dans la feuille" & vbNewLine & vbNewLine & _
                    "BD_Clients...", vbExclamation, "Recherche dans BD_Clients " & dynamicRange.Address
    End If
    
    'Libérer la mémoire
    Set dynamicRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFunctions:Fn_GetID_From_Client_Name", "", startTime)

End Function

Function Fn_Cell_From_BD_Client(nomClient As String, ByRef colNumberSearch As Integer, ByRef colNumberData As Integer) As String '2025-01-12 @ 08:12

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFunctions:Fn_Cell_From_BD_Client", nomClient, 0)
    
    Dim ws As Worksheet: Set ws = wshBD_Clients
    
    On Error Resume Next
    Dim dynamicRange As Range: Set dynamicRange = ws.Range("dnrClients_All")
    On Error GoTo 0

    If ws Is Nothing Or dynamicRange Is Nothing Then
        MsgBox "La feuille 'Clients' ou le DynamicRange 'dnrClients_All' n'a pas été trouvé!", _
            vbExclamation
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly, requires EXACT match (5th parameter = 0 ) - 2025-01-12 @ 14:49
    Dim result As Variant
    result = Application.WorksheetFunction.XLookup(nomClient, _
                                                   dynamicRange.Columns(colNumberSearch), _
                                                   dynamicRange.Columns(colNumberData), _
                                                   "Not Found", _
                                                   0, _
                                                   1)
    If result <> "Not Found" Then
        Fn_Cell_From_BD_Client = result
    Else
        MsgBox "Impossible de retrouver le nom du client dans la feuille" & vbNewLine & vbNewLine & _
                    "BD_Clients...", vbExclamation, "Fn_Cell_From_BD_Client"
    End If
    
    'Libérer la mémoire
    Set dynamicRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFunctions:Fn_Cell_From_BD_Client", "", startTime)

End Function

Function Fn_GetID_From_Fourn_Name(nomFournisseur As String) '2024-07-03 @ 16:13

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFunctions:Fn_GetID_From_Fourn_Name", nomFournisseur, 0)
    
    Dim ws As Worksheet: Set ws = wshBD_Fournisseurs
    
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
        Fn_GetID_From_Fourn_Name = result
    Else
        Fn_GetID_From_Fourn_Name = 0
    End If
    
    'Libérer la mémoire
    Set dynamicRange = Nothing
    Set ws = Nothing

    Call Log_Record("modFunctions:Fn_GetID_From_Fourn_Name", "", startTime)

End Function

Function Fn_Get_Prenom_From_Initials(i As String)

    Dim cell As Range
    
    For Each cell In wshAdmin.Range("dnrProf_All")
        If cell.Value2 = i Then
            Fn_Get_Prenom_From_Initials = cell.offset(0, 2).value
            Exit Function
        End If
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Function Fn_Get_Nom_From_Initials(i As String)

    Dim cell As Range
    
    For Each cell In wshAdmin.Range("dnrProf_All")
        If cell.Value2 = i Then
            Fn_Get_Nom_From_Initials = cell.offset(0, 3).value
            Exit Function
        End If
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Function Fn_Get_Value_From_UniqueID(ws As Worksheet, uniqueID As String, keyColumn As Integer, returnColumn As Integer) As Variant

    'Définir la dernière ligne utilisée de la feuille
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, keyColumn).End(xlUp).row
    
    'Définir la plage de recherche (toute la colonne de la clé)
    Dim searchRange As Range
    Set searchRange = ws.Range(ws.Cells(1, keyColumn), ws.Cells(lastRow, keyColumn))
    
    'Rechercher la clé dans la colonne spécifiée
    Dim foundCell As Range
    Set foundCell = searchRange.Find(What:=uniqueID, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Si on a trouvé 'uniqueID', retourner la valeur de la colonne de retour
    If Not foundCell Is Nothing Then
        Fn_Get_Value_From_UniqueID = ws.Cells(foundCell.row, returnColumn).value
    Else
        'Si l'on a pas trouvée, retourner une valeur d'erreur ou un message
        Fn_Get_Value_From_UniqueID = "uniqueID introuvable"
    End If
    
    'Libérer la mémoire
    Set foundCell = Nothing
    Set searchRange = Nothing
    Set ws = Nothing
    
End Function

Function Fn_Find_Data_In_A_Range(r As Range, cs As Long, ss As String, cr As Long) As Variant() '2024-03-29 @ 05:39
    
    'This function is used to retrieve information from in a range(r) at column (cs) the value of (ss)
    'If found, it returns an array, with the cell address(1), the row(2) and the value of column cr(3)
    'Otherwise it return an empty array
    '2024-03-09 - First version
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFunctions:Fn_Find_Data_In_A_Range", "", 0)
    
    Dim foundInfo(1 To 3) As Variant 'Cell Address, Row, Value
    
    'Search for the string in a given range (r) at the column specified (cs)
    Dim foundCell As Range: Set foundCell = r.Columns(cs).Find(What:=ss, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the string was found
    If Not foundCell Is Nothing Then
        'With the foundCell get the the address, the row number and the value
        foundInfo(1) = foundCell.Address
        foundInfo(2) = foundCell.row
        foundInfo(3) = foundCell.offset(0, cr - cs).value 'Return Column - Searching column
        Fn_Find_Data_In_A_Range = foundInfo 'foundInfo is an array
    Else
        Fn_Find_Data_In_A_Range = foundInfo 'foundInfo is an array
    End If
    
    'Libérer la mémoire
    Set foundCell = Nothing

    Call Log_Record("modFunctions:Fn_Find_Data_In_A_Range", "", startTime)

End Function

Function Fn_Valider_Courriel(ByVal courriel As String) As Boolean
    
    Fn_Valider_Courriel = False
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    'Définir le pattern pour l'expression régulière
    regex.Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    regex.IgnoreCase = True
    regex.Global = False
    
    'Last chance to accept a invalid email address...
    If regex.Test(courriel) = False Then
        Dim msgValue As VbMsgBoxResult
        msgValue = MsgBox("'" & courriel & "'" & vbNewLine & vbNewLine & _
                            "N'est pas structurée selon les standards..." & vbNewLine & vbNewLine & _
                            "Désirez-vous quand même conserver cette adresse ?", _
                            vbYesNo + vbInformation, "Struture de courriel non standard")
        If msgValue = vbYes Then
            Fn_Valider_Courriel = True
        Else
            Fn_Valider_Courriel = False
        End If
    Else
        Fn_Valider_Courriel = True
    End If
    
    'Libérer la mémoire
    Set regex = Nothing
    
End Function

Function Fn_Verify_And_Delete_Rows_If_Value_Is_Found(valueToFind As Variant, hono As Double) As String '2024-07-18 @ 16:32
    
    'Define the worksheet
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    
    'Define the range to search in (Column 1)
    Dim searchRange As Range: Set searchRange = ws.Columns(2)
    
    'Search for the first occurrence of the value
    Dim cell As Range
    Set cell = searchRange.Find(What:=valueToFind, _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole)
    
    'Check if the value is found
    Dim firstAddress As String
    Dim rowsToDelete As Collection: Set rowsToDelete = New Collection

    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Fn_Verify_And_Delete_Rows_If_Value_Is_Found = firstAddress
        
        'Loop to collect all rows with the value
        Do
            rowsToDelete.Add cell.row
            Set cell = searchRange.FindNext(cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        
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
                    Fn_Verify_And_Delete_Rows_If_Value_Is_Found = "REMPLACER"
                End If
                If reponse = vbCancel Then
                    Fn_Verify_And_Delete_Rows_If_Value_Is_Found = "SUPPRIMER"
                End If
                
                'Delete all collected rows from wshFAC_Projets_Détails (locally)
                Dim i As Long
                For i = rowsToDelete.count To 1 Step -1
                    ws.Rows(rowsToDelete(i)).Delete
                Next i
                
                'Update rows from MASTER file (details)
                Dim destinationFileName As String, destinationTab As String
                destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                                      "GCF_BD_MASTER.xlsx"
                destinationTab = "FAC_Projets_Détails$"
                
                Dim columnName As String
                columnName = "NomClient"
                Call Soft_Delete_If_Value_Is_Found_In_Master_Details(destinationFileName, _
                                                                     destinationTab, _
                                                                     columnName, _
                                                                     valueToFind)
                                                                     
                'Update row from MASTER file (entête)
                destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                                      "GCF_BD_MASTER.xlsx"
                destinationTab = "FAC_Projets_Entête$"
                Call Soft_Delete_If_Value_Is_Found_In_Master_Entete(destinationFileName, _
                                                                    destinationTab, _
                                                                    columnName, _
                                                                    valueToFind) '2024-07-19 @ 15:31
            Case vbNo
                Fn_Verify_And_Delete_Rows_If_Value_Is_Found = "RIEN_CHANGER"
        End Select
    Else
        Fn_Verify_And_Delete_Rows_If_Value_Is_Found = "REMPLACER"
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set rowsToDelete = Nothing
    Set searchRange = Nothing
    Set ws = Nothing
    
End Function

Function Fn_GetCheckBoxPosition(chkBox As OLEObject) As String

    'Get the cell that contains the top-left corner of the CheckBox
    Fn_GetCheckBoxPosition = chkBox.TopLeftCell.Address
    
End Function

Function Fn_Get_Column_Type(col As Range) As String

    Dim cell As Range
    Dim dataType As String
    Dim cellValue As Variant
    
    dataType = "Empty" ' Default type if no data found
    
    ' Loop through cells in the first few rows to determine the data type
    For Each cell In col.Cells
        cellValue = cell.value
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
    
    Fn_Get_Column_Type = dataType
    
    'Libérer la mémoire
    Set cell = Nothing
    
End Function

Public Function Fn_GetGL_Code_From_GL_Description(glDescr As String) 'XLOOKUP - 2024-01-09 @ 09:19

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFunctions:Fn_GetGL_Code_From_GL_Description", glDescr, 0)
    
    Dim ws As Worksheet: Set ws = wshAdmin
    
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
    
    Call Log_Record("     modFunctions:Fn_GetGL_Code_From_GL_Description - " & result, -1)
    
    If result <> "Not Found" Then
        Fn_GetGL_Code_From_GL_Description = result
    Else
        MsgBox "Impossible de retrouver la valeur dans la première colonne", vbExclamation
    End If

    'Libérer la mémoire
    Set dynamicRange = Nothing
    Set ws = Nothing

    Call Log_Record("modFunctions:Fn_GetGL_Code_From_GL_Description", "", startTime)

End Function

Function Fn_Get_GL_Account_Balance(glCode As String, dateSolde As Date) As Currency '2024-11-18 @ 06:41
    
    Fn_Get_GL_Account_Balance = 0
    
    'AdvancedFilter GL_Trans with FromDate to ToDate, returns rngResult
    Dim rngResult As Range
    Call GL_Get_Account_Trans_AF(glCode, #7/31/2024#, dateSolde, rngResult)
    
    'Méthode plus rapide pour obtenir une somme
    Fn_Get_GL_Account_Balance = Application.WorksheetFunction.Sum(rngResult.Columns(7)) _
                                           - Application.WorksheetFunction.Sum(rngResult.Columns(8))

End Function

Function Fn_Get_GL_Month_Trans_Total(glCode As String, dateFinMois As Date) As Double '2025-02-07 @ 13:46
    
    Fn_Get_GL_Month_Trans_Total = 0
    
    Dim dateDebutMois As Date
    dateDebutMois = DateSerial(year(dateFinMois), month(dateFinMois), 1)
    
    'AdvancedFilter GL_Trans with FromDate to ToDate, returns rngResult
    Dim rngResult As Range
    Call GL_Get_Account_Trans_AF(glCode, dateDebutMois, dateFinMois, rngResult)
    
    'Méthode plus rapide pour obtenir une somme
    Fn_Get_GL_Month_Trans_Total = Application.WorksheetFunction.Sum(rngResult.Columns(7)) _
                                           - Application.WorksheetFunction.Sum(rngResult.Columns(8))

End Function

Function Fn_ObtenirTECFacturésPourFacture(invNo As String) As Variant

    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    Dim lastUsedRow As Long
    lastUsedRow = wsTEC.Cells(wsTEC.Rows.count, 1).End(xlUp).row '2024-08-18 @ 06:37
    
    Dim resultArr() As Variant
    ReDim resultArr(1 To 1000)
    
    Dim rowCount As Long
    Dim i As Long
    For i = 3 To lastUsedRow
        If wsTEC.Cells(i, 16).value = invNo And UCase(wsTEC.Cells(i, 14).value) <> "VRAI" Then
            rowCount = rowCount + 1
            resultArr(rowCount) = i
        End If
    Next i
    
    If rowCount > 0 Then
        ReDim Preserve resultArr(1 To rowCount)
    End If
    
    If rowCount = 0 Then
        Fn_ObtenirTECFacturésPourFacture = Array()
    Else
        Fn_ObtenirTECFacturésPourFacture = resultArr
    End If
    
    'Libérer la mémoire
    Set wsTEC = Nothing
    
End Function

Function Fn_Get_TEC_Total_Invoice_AF(invNo As String, t As String) As Currency

    'Le type (t) est "Heures" -OU- "Dollars", selon le type le total des Heures ou des Dollars
    
    Fn_Get_TEC_Total_Invoice_AF = 0
    
    Dim ws As Worksheet: Set ws = wshFAC_Détails
    
    'Effacer les données de la dernière utilisation
    ws.Range("I6:I10").ClearContents
    ws.Range("I6").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_FAC_Détails[#All]")
    ws.Range("I7").value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("I2:I3")
    ws.Range("I3").value = invNo
    ws.Range("I8").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("K1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("K2:N2")
    ws.Range("I9").value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "K").End(xlUp).row
    ws.Range("I10").value = lastUsedRow - 2 & " lignes"
    
    'Aucun tri nécessaire (besoins)
    If lastUsedRow > 2 Then
        Dim i As Long
        For i = 3 To lastUsedRow
            If InStr(ws.Cells(i, 11), "*** - [Sommaire des TEC] pour la facture - ") = 1 Then
                If t = "Heures" Then
                    Fn_Get_TEC_Total_Invoice_AF = Fn_Get_TEC_Total_Invoice_AF + ws.Cells(i, "L")
                Else
                    Fn_Get_TEC_Total_Invoice_AF = Fn_Get_TEC_Total_Invoice_AF + ws.Cells(i, "N")
                End If
            End If
        Next i
    End If
    
    'Force un arrondissement à 2 décimales
    Fn_Get_TEC_Total_Invoice_AF = Round(Fn_Get_TEC_Total_Invoice_AF, 2)
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
End Function

Public Function Fn_Find_Row_Number_TECID(ByVal uniqueID As Variant, ByVal lookupRange As Range) As Long '2024-08-10 @ 05:41
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFunctions:Fn_Find_Row_Number_TECID", "", 0)
    
    On Error Resume Next
        Dim cell As Range
        Set cell = lookupRange.Find(What:=uniqueID, LookIn:=xlValues, LookAt:=xlWhole)
        If Not cell Is Nothing Then
            Fn_Find_Row_Number_TECID = cell.row
            Call Log_Record("modFunctions:Fn_Find_Row_Number_TECID" & " - Row # = " & Fn_Find_Row_Number_TECID, -1)
        Else
            Fn_Find_Row_Number_TECID = -1 'Not found
            Call Log_Record("modFunctions:Fn_Find_Row_Number_TECID" & " - TECID = WAS NOT FOUND...", -1)
        End If
    On Error GoTo 0
    
    'Libérer la mémoire
    Set cell = Nothing
    
    Call Log_Record("modFunctions:Fn_Find_Row_Number_TECID", "", startTime)
    
End Function

Function Fn_Get_Bucket_For_Aging(age As Long, days1 As Long, days2 As Long, days3 As Long, days4 As Long)

    Select Case age
        Case Is < days1
            Fn_Get_Bucket_For_Aging = 0
        Case Is < days2
            Fn_Get_Bucket_For_Aging = 1
        Case Is < days3
            Fn_Get_Bucket_For_Aging = 2
        Case Is < days4
            Fn_Get_Bucket_For_Aging = 3
        Case Else
            Fn_Get_Bucket_For_Aging = 4
    End Select
    
End Function

Function Fn_Get_Invoice_Total_Payments_AF(invNo As String)

    Fn_Get_Invoice_Total_Payments_AF = 0
    
    Dim ws As Worksheet: Set ws = wshENC_Détails
    
    'Effacer les données de la dernière utilisation
    ws.Range("H6:H10").ClearContents
    ws.Range("H6").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_ENC_Détails[#All]")
    ws.Range("H7").value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("H2:H3")
    ws.Range("H3").value = invNo
    ws.Range("H8").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("J1").CurrentRegion
    rngResult.offset(3, 0).Clear
    Set rngResult = ws.Range("J3:N3")
    ws.Range("H9").value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "J").End(xlUp).row
    ws.Range("H10").value = lastUsedRow - 3 & " lignes"
    
    'Il n'est pas nécessaire de trier les résultats
    If lastUsedRow > 3 Then
        Set rngResult = ws.Range("J4:N" & lastUsedRow)
        Fn_Get_Invoice_Total_Payments_AF = Application.WorksheetFunction.Sum(rngResult.Columns(5))
    End If

    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
End Function

Function Fn_Get_A_Cell_From_A_Worksheet(feuille, cle As String, cleCol As Integer, retourCol As Integer) As String '2025-02-22 @ 06:56

    'Définir la feuille pour la recherche
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(feuille)
    
    'Définir la plage pour le résultat
    Dim resultat As Range
    
    'Utilisation de la méthode Find pour rechercher dans la première colonne
    Set resultat = ws.Columns(cleCol).Find(What:=cle, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not resultat Is Nothing Then
        Fn_Get_A_Cell_From_A_Worksheet = ws.Cells(resultat.row, retourCol)
    Else
        Fn_Get_A_Cell_From_A_Worksheet = ""
    End If
    
    'Libérer la mémoire
    Set resultat = Nothing
    Set ws = Nothing

End Function

'Function Fn_GetCellFromComptesClients_ACO(invNo As String, col As Integer) As String
'
'    'Définir la feuille pour la recherche
'    Dim ws As Worksheet
'    Set ws = wshFAC_Comptes_Clients
'
'    'Définir la plage pour le résultat
'    Dim foundCell As Range
'
'    'Utilisation de la méthode Find pour rechercher dans la première colonne
'    Set foundCell = ws.Columns(1).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
'
'    If Not foundCell Is Nothing Then
'        Fn_GetCellFromComptesClients = ws.Cells(foundCell.row, col)
'    Else
'        Fn_GetCellFromComptesClients = ""
'    End If
'
'    'Libérer la mémoire
'    Set foundCell = Nothing
'    Set ws = Nothing
'
'End Function
'
Function Fn_Validate_And_Get_A_Cell(ws As Worksheet, search As String, searchCol As Long, returnCol As Long) As Variant

    Dim foundCell As Range
    
    'Utilisation de la méthode Find pour rechercher dans la première colonne
    Set foundCell = ws.Columns(searchCol).Find(What:=search, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        Fn_Validate_And_Get_A_Cell = ws.Cells(foundCell.row, returnCol)
    Else
        Fn_Validate_And_Get_A_Cell = ""
    End If
    
    'Libérer la mémoire
    Set foundCell = Nothing

End Function

Function Fn_Validate_Client_Number(clientCode As String) As Boolean '2024-10-26 @ 18:30

    '2024-08-14 @ 10:17 - Verify that a client exists, based on clientCode
    
    Fn_Validate_Client_Number = False
    
    Dim lastUsedRow As Long
    lastUsedRow = wshBD_Clients.Cells(wshBD_Clients.Rows.count, "B").End(xlUp).row
    Dim rngToSearch As Range
    Set rngToSearch = wshBD_Clients.Range("B1:B" & lastUsedRow)
    
    'Search for the string in a given range (r) at the column specified (cs)
    Dim rngFound As Range
    Set rngFound = rngToSearch.Find(What:=clientCode, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        Fn_Validate_Client_Number = True
    Else
        Fn_Validate_Client_Number = False
    End If

    'Clean-up - 2024-08-14 @ 10:15
    Set rngFound = Nothing
    Set rngToSearch = Nothing
    
End Function

Function Fn_ValiderCourriel(ByVal adresses As String) As Boolean '2024-10-26 @ 14:30
    
    'Validation de 0 à n adresses courriel
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    'Initialisation de l'expression régulière pour valider une adresse courriel
    With regex
        .Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
        .IgnoreCase = True
        .Global = False
    End With
    
    'Diviser le paremètre (courriel) en adresses individuelles
    Dim adressesArray() As String
    adressesArray = Split(adresses, ";")
    
    ' Vérifier chaque adresse
    Dim adresse As Variant
    For Each adresse In adressesArray
        adresse = Trim(adresse)
        'Passer si l'adresse est vide (Aucune adresse est aussi permis)
        If adresse <> "" Then
            'Si l'adresse ne correspond pas au pattern, renvoyer Faux
            If Not regex.Test(adresse) Then
                Fn_ValiderCourriel = False
                Exit Function
            End If
        End If
    Next adresse
    
    ' Toutes les adresses sont valides
    Fn_ValiderCourriel = True
    
    'Nettoyer la mémoire
    Set adresse = Nothing
    Set regex = Nothing
    
End Function

Function Fn_ValidateDaySpecificMonth(d As Long, m As Long, y As Long) As Boolean
    'Returns TRUE or FALSE if d, m and y combined are VALID values
    
    Fn_ValidateDaySpecificMonth = False
    
    Dim isLeapYear As Boolean
    If y Mod 4 = 0 And (y Mod 100 <> 0 Or y Mod 400 = 0) Then
        isLeapYear = True
    Else
        isLeapYear = False
    End If
    
    'Last day of each month (0 to 11)
    Dim mdpm As Variant
    mdpm = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    If isLeapYear Then mdpm(1) = 29 'Adjust February for Leap Year
    
    If m < 1 Or m > 12 Or _
       d > mdpm(m - 1) Or _
       Abs(year(Date) - y) > 75 Then
            Exit Function
    Else
        Fn_ValidateDaySpecificMonth = True
    End If

End Function

Function Fn_Check_Server_Access(serverPath) As Boolean '2024-09-24 @ 17:14

    DoEvents
    
    Fn_Check_Server_Access = False
    
    'Créer un FileSystemObject
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'Vérifier si le fichier existe
    Dim folderExists As Boolean
    folderExists = FSO.folderExists(serverPath)
    
    If folderExists Then
        Fn_Check_Server_Access = True
    End If
    
    'Libérer la mémoire
    Set FSO = Nothing
    
End Function

Function Fn_Is_Server_Available() As Boolean

    DoEvents
    
    On Error Resume Next
    'Tester l'existence d'un fichier ou d'un répertoire sur le lecteur P:
    If Dir("P:\", vbDirectory) <> "" Then
        Fn_Is_Server_Available = True
    Else
        Fn_Is_Server_Available = False
    End If
    On Error GoTo 0
    
End Function

Function Fn_Complete_Date(dateInput As String, joursArriere As Integer, joursFutur As Integer) As Variant
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFunctions:Fn_Complete_Date", dateInput & ", " & joursArriere & ", " & joursFutur, 0)
    
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
    dateInput = Replace(Replace(Replace(dateInput, "/", "-"), ".", "-"), " ", "")
    Dim parts() As String
    parts = Split(Replace(dateInput, "-01-1900", ""), "-")
    
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
    If Fn_ValidateDaySpecificMonth(dayPart, monthPart, yearPart) = False Then
        GoTo Invalid_Date
    End If
    
    'Construct the full date
    parsedDate = DateSerial(yearPart, monthPart, dayPart)
    Dim joursEcart As Integer
    joursEcart = parsedDate - Date
    If joursEcart < 0 And Abs(joursEcart) > joursArriere Then
        MsgBox "Cette date NE RESPECTE PAS les paramètres de date établis" & vbNewLine & vbNewLine & _
                    "La date minimale est '" & Format$(Date - joursArriere, wshAdmin.Range("B1").value) & "'", _
                    vbCritical, "La date saisie est hors-norme - (Du " & _
                        Format$(Date - joursArriere, wshAdmin.Range("B1").value) & " au " & Format$(Date + joursFutur, wshAdmin.Range("B1").value) & ")"
        GoTo Invalid_Date
    End If
    If joursEcart > 0 And joursEcart > joursFutur Then
        MsgBox "Cette date NE RESPECTE PAS les paramètres de date établis" & vbNewLine & vbNewLine & _
                    "La date maximale est '" & Format$(Date + joursFutur, wshAdmin.Range("B1").value) & "'", _
                    vbCritical, "La date saisie est hors-norme - (Du " & _
                    Format$(Date - joursArriere, wshAdmin.Range("B1").value) & " au " & Format$(Date + joursFutur, wshAdmin.Range("B1").value) & ")"
        GoTo Invalid_Date
    End If
   
    'Return a VALID date
    Fn_Complete_Date = parsedDate
    
    Call Log_Record("modFunctions:Fn_Complete_Date", "", startTime)

    Exit Function

Invalid_Date:

    Fn_Complete_Date = "Invalid Date"
    
    Call Log_Record("modFunctions:Fn_Complete_Date", "", startTime)
    
End Function

Function Fn_Sort_Dictionary_By_Keys(dict As Object, Optional descending As Boolean = False) As Variant '2024-10-02 @ 12:02
    
    'Sort a dictionary by its keys and return keys in an array
    Dim keys() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    
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
    
    Fn_Sort_Dictionary_By_Keys = keys
    
    'Libérer la mémoire
    Set key = Nothing
    
End Function

Function Fn_Sort_Dictionary_By_Value(dict As Object, Optional descending As Boolean = False) As Variant '2024-07-11 @ 15:16
    
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
    
    Fn_Sort_Dictionary_By_Value = keys
    
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
        Fn_Strip_Contact_From_Client_Name = Trim(Left(cn, posOSB - 1) & Mid(cn, posCSB + 1))
    Else
        Fn_Strip_Contact_From_Client_Name = Trim(Mid(cn, posCSB + 1))
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
    If ufSaisieHeures.cmbProfessionnel.value = "" Then
        MsgBox prompt:="Le professionnel est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        ufSaisieHeures.cmbProfessionnel.SetFocus
        Exit Function
    End If

    'Date de la charge ?
    If ufSaisieHeures.txtDate.value = "" Or IsDate(ufSaisieHeures.txtDate.value) = False Then
        MsgBox prompt:="La date est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        ufSaisieHeures.txtDate.SetFocus
        Exit Function
    End If

    'Nom du client & code de client ?
    If ufSaisieHeures.txtClient.value = "" Or ufSaisieHeures.txtClientID = "" Then
        MsgBox prompt:="Le client et son code sont OBLIGATOIRES !" & vbNewLine & vbNewLine & _
                       "Code de client = '" & ufSaisieHeures.txtClientID & "'" & vbNewLine & vbNewLine & _
                       "Nom du client = '" & ufSaisieHeures.txtClient.value & "'", _
               Title:="Vérifications essentielles des données du client", _
               Buttons:=vbCritical
        ufSaisieHeures.txtClient.SetFocus
        Exit Function
    End If
    
    'Heures valides ?
    If ufSaisieHeures.txtHeures.value = "" Or IsNumeric(ufSaisieHeures.txtHeures.value) = False Then
        MsgBox prompt:="Le nombre d'heures est OBLIGATOIRE !", _
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
            If rowRange.Cells(1, 1).value = profID Then
                If CDate(dte) >= CDate(rowRange.Cells(1, 2).value) Then
                    Fn_Get_Hourly_Rate = rowRange.Cells(1, 3).value
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

Function Fn_Get_Invoice_Type(invNo As String) As String '2024-08-17 @ 06:55

    'Return the Type of invoice - 'C' for confirmed, 'AC' to be confirmed
    
    Dim lastUsedRow As Long
    lastUsedRow = wshFAC_Entête.Cells(wshFAC_Entête.Rows.count, 1).End(xlUp).row
    Dim rngToSearch As Range
    Set rngToSearch = wshFAC_Entête.Range("A1:A" & lastUsedRow)
    
    'Find the invNo into rngToSearch
    Dim rngFound As Range
    Set rngFound = rngToSearch.Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFound Is Nothing Then
        Fn_Get_Invoice_Type = rngFound.offset(0, 2).value
    Else
        Fn_Get_Invoice_Type = "C"
    End If

    'Clean-up - 2024-08-17 @ 06:55
    Set rngFound = Nothing
    Set rngToSearch = Nothing
    
End Function

Public Function Fn_Get_Tax_Rate(d As Date, taxType As String) As Double

    Dim row As Long
    Dim rate As Double
    With wshAdmin
        For row = 18 To 11 Step -1
            If .Range("L" & row).value = taxType Then
                If d >= .Range("M" & row).value Then
                    rate = .Range("N" & row).value
                    Exit For
                End If
            End If
        Next row
    End With
    
    Fn_Get_Tax_Rate = rate
    
End Function

Public Function Fn_Is_Client_Facturable(ByVal clientID As String) As Boolean

    Fn_Is_Client_Facturable = False
    
    'Les clients NON FACTURABLES sont compris entre 1 et 99
    If Len(clientID) > 2 Then
        Fn_Is_Client_Facturable = True
    End If
        
End Function

Function Fn_Is_Date_Valide(d As String) As Boolean

    Fn_Is_Date_Valide = False
    If d = "" Or IsDate(d) = False Then
        MsgBox "Une date d'écriture est obligatoire." & vbNewLine & vbNewLine & _
            "Veuillez saisir une date valide!", vbCritical, "Date Invalide"
    Else
        Fn_Is_Date_Valide = True
    End If

End Function

Function Fn_Get_Windows_Username() As String 'Function to retrieve the Windows username using the API

    Dim buffer As String * 255
    Dim size As Long: size = 255
    
    If GetUserName(buffer, size) Then
        Fn_Get_Windows_Username = Left$(buffer, size - 1)
    Else
        Fn_Get_Windows_Username = "Unknown"
    End If
    
End Function

Function Fn_Invoice_Is_Confirmed(invNo As String) As Boolean

    Fn_Invoice_Is_Confirmed = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Entête

    'Utilisation de FIND pour trouver la cellule contenant la valeur recherchée dans la colonne A
    Dim foundCell As Range
    Set foundCell = ws.Range("A:A").Find(What:=CStr(invNo), LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        If foundCell.offset(0, 2).value = "C" Then
            Fn_Invoice_Is_Confirmed = True
        End If
    Else
        Fn_Invoice_Is_Confirmed = False
    End If

    'Libérer la mémoire
    Set foundCell = Nothing
    Set ws = Nothing

End Function

Function Fn_Is_Ecriture_Balance() As Boolean

    Fn_Is_Ecriture_Balance = False
    If wshGL_EJ.Range("H26").value <> wshGL_EJ.Range("I26").value Then
        MsgBox "Votre écriture ne balance pas." & vbNewLine & vbNewLine & _
            "Débits = " & wshGL_EJ.Range("H26").value & " et Crédits = " & wshGL_EJ.Range("I26").value & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée.", vbCritical, "Veuillez vérifier votre écriture!"
    Else
        Fn_Is_Ecriture_Balance = True
    End If

End Function

Function Fn_Is_Debours_Balance() As Boolean

    Fn_Is_Debours_Balance = False
    If wshDEB_Saisie.Range("O6").value <> wshDEB_Saisie.Range("I26").value Then
        MsgBox "Votre transaction ne balance pas." & vbNewLine & vbNewLine & _
            "Total saisi = " & Format$(wshDEB_Saisie.Range("O6").value, "#,##0.00 $") _
            & " vs. Ventilation = " & Format$(wshDEB_Saisie.Range("I26").value, "#,##0.00 $") _
            & vbNewLine & vbNewLine & "Elle n'est donc pas reportée.", _
            vbCritical, "Veuillez vérifier votre écriture!"
    Else
        Fn_Is_Debours_Balance = True
    End If

End Function

Function Fn_Is_JE_Valid(rmax As Long) As Boolean

    Fn_Is_JE_Valid = True 'Optimist
    If rmax <= 9 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        Fn_Is_JE_Valid = False
    End If
    
    Dim i As Long
    For i = 9 To rmax
        If wshGL_EJ.Range("E" & i).value <> "" Then
            If wshGL_EJ.Range("H" & i).value = "" And wshGL_EJ.Range("I" & i).value = "" Then
                MsgBox "Il existe une ligne avec un compte, sans montant !"
                Fn_Is_JE_Valid = False
            End If
        End If
    Next i

End Function

Function Fn_Is_Deb_Saisie_Valid(rmax As Long) As Boolean

    Fn_Is_Deb_Saisie_Valid = True 'Optimist
    If rmax < 9 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        Fn_Is_Deb_Saisie_Valid = False
    End If
    
    Dim i As Long
    For i = 9 To rmax
        If wshDEB_Saisie.Range("E" & i).value <> "" Then
            If wshDEB_Saisie.Range("N" & i).value = "" Then
                MsgBox "Il existe une ligne avec un compte, sans montant !"
                Fn_Is_Deb_Saisie_Valid = False
            End If
        End If
    Next i

End Function

Public Function Fn_Pad_A_String(s As String, fillCaracter As String, length As Long, leftOrRight As String) As String

    Dim paddedString As String
    Dim charactersNeeded As Long
    
    charactersNeeded = length - Len(s)
    
    If charactersNeeded > 0 Then
        If leftOrRight = "R" Then
            paddedString = s & String(charactersNeeded, fillCaracter)
        Else
            paddedString = String(charactersNeeded, fillCaracter) & s
        End If
    Else
        paddedString = s
    End If

    Fn_Pad_A_String = paddedString
        
End Function

Function Fn_Get_Next_Invoice_Number() As String '2024-09-17 @ 14:00

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim strLastInvoice As String
    strLastInvoice = ws.Cells(lastUsedRow, 1).value
    If strLastInvoice <> "" Then
        strLastInvoice = Right(strLastInvoice, Len(strLastInvoice) - 3)
    Else
        MsgBox "Problème avec les dernières lignes de la" & _
                vbNewLine & vbNewLine & "feuille 'wshFAC_Entête'" & _
                vbNewLine & vbNewLine & "Veuillez contacter le développeur", _
                vbOKOnly, "Structure invalide dans 'wshFAC_Entête'"
    End If
    Fn_Get_Next_Invoice_Number = strLastInvoice + 1

    'Libérer la mémoire
    Set ws = Nothing
    
End Function

Function Fn_Get_GL_Account_Opening_Balance_AF(glNo As String, d As Date) As Double

    'Using AdvancedFilter # 1 in wshGL_Trans
    
    Fn_Get_GL_Account_Opening_Balance_AF = 0
    
    Dim ws As Worksheet: Set ws = wshGL_Trans
    
    Application.EnableEvents = False
    
    'Effacer les données de la dernière utilisation
    ws.Range("M6:M10").ClearContents
    ws.Range("M6").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_GL_Trans[#All]")
    ws.Range("M7").value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("L2:N3")
    ws.Range("L3").FormulaR1C1 = glNo
    ws.Range("M3").FormulaR1C1 = ">=" & CLng(#7/31/2024#)
    ws.Range("N3").FormulaR1C1 = "<" & CLng(d)
    ws.Range("M8").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("P1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("P1:Y1")
    ws.Range("M9").value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "P").End(xlUp).row
    ws.Range("M10").value = lastUsedRow - 1 & " lignes"
    
    Application.EnableEvents = True
    
    'Pas de tri nécessaire pour calculer le solde
    If lastUsedRow < 2 Then
        Exit Function
    End If
        
    'Méthode plus rapide pour obtenir une somme
    Set rngResult = ws.Range("P2:Y" & lastUsedRow)
    Fn_Get_GL_Account_Opening_Balance_AF = Application.WorksheetFunction.Sum(rngResult.Columns(7)) _
                                           - Application.WorksheetFunction.Sum(rngResult.Columns(8))
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
End Function

Function Fn_Get_Plan_Comptable(nbCol As Long) As Variant '2024-06-07 @ 07:31

    Debug.Assert nbCol >= 1 And nbCol <= 4 '2024-07-31 @ 19:26
    
    'Reference the named range
    Dim planComptable As Range: Set planComptable = wshAdmin.Range("dnrPlanComptable_All")
    
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
    
    Fn_Get_Plan_Comptable = arr
    
    'Libérer la mémoire
    Set planComptable = Nothing
    Set row = Nothing
    Set rowRange = Nothing
    
End Function

Function Fn_Get_Client_Name(cc As String) As String

    Dim ws As Worksheet
    Dim foundCell As Range
    
    Set ws = wshBD_Clients
    
    'Recherche le code de client dans la colonne B
    Set foundCell = ws.Columns("B").Find(What:=cc, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        'Si trouvé, retourner le nom du client correspondant, 1 colonne à gauche
        Fn_Get_Client_Name = foundCell.offset(0, -1).value
    Else
        Fn_Get_Client_Name = "Client non trouvé (invalide)"
    End If
    
    'Libérer la mémoire
    Set foundCell = Nothing
    Set ws = Nothing
    
End Function

Function Fn_Rechercher_Client_Par_ID(codeClient As String, ws As Worksheet) As Variant

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
        clientData = ws.Rows(ligneTrouvee).value
        
        'Retourner les données du client (ligne entière)
        Fn_Rechercher_Client_Par_ID = clientData
    Else
        'Si le client n'est pas trouvé, retourner une valeur vide ou une erreur
        Fn_Rechercher_Client_Par_ID = CVErr(xlErrNA) 'Retourne #N/A pour indiquer que le client n'est pas trouvé
    End If
    
    'Libérer la mémoire
    Set foundCells = Nothing
    Set rangeID = Nothing
    
End Function

Function Fn_Remove_All_Accents(ByVal Text As String) As String

    'Liste des caractères accentués et leurs équivalents sans accents
    Dim AccChars As String
    AccChars = "ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
    Dim RegChars As String
    RegChars = "AAAAAACEEEEIIIINOOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"

    'Remplacer les accents par des caractères non accentués
    Dim i As Long
    For i = 1 To Len(AccChars)
        Text = Replace(Text, Mid(AccChars, i, 1), Mid(RegChars, i, 1))
    Next i

    Fn_Remove_All_Accents = Text
    
End Function

Public Function Fn_Get_Current_Region(ByVal dataRange As Range, Optional HeaderSize As Long = 1) As Range

    Set Fn_Get_Current_Region = dataRange.CurrentRegion
    If HeaderSize > 0 Then
        With Fn_Get_Current_Region
            'Remove the header
            Set Fn_Get_Current_Region = .offset(HeaderSize).Resize(.Rows.count - HeaderSize)
            Debug.Print "#060 - " & Fn_Get_Current_Region.Address
        End With
    End If
    
    'Libérer la mémoire
    Set Fn_Get_Current_Region = Nothing
    
End Function

Public Function Fn_Convert_Value_Boolean_To_Text(val As Boolean) As String

    Select Case val
        Case 0, "False", "Faux" 'False
            Fn_Convert_Value_Boolean_To_Text = "FAUX"
        Case -1, "True", "Vrai" 'True"
            Fn_Convert_Value_Boolean_To_Text = "VRAI"
        Case "VRAI", "FAUX"
            
        Case Else
            MsgBox val & " est une valeur INVALIDE !"
    End Select

End Function

Function Fn_Is_String_Valid(searchString As String, rng As Range) As Boolean

    On Error Resume Next
    Fn_Is_String_Valid = Not IsError(Application.Match(searchString, rng, 0))
    On Error GoTo 0
    
End Function

Function Fn_Count_Char_Occurrences(ByVal inputString As String, ByVal charToCount As String) As Long
    
    'Ensure charToCount is a single character
    If Len(charToCount) <> 1 Or Len(inputString) = 0 Then
        Fn_Count_Char_Occurrences = -1 ' Return -1 for invalid input
        Exit Function
    End If
    
    'Loop through each character in the string
    Dim i As Long, count As Long
    For i = 1 To Len(inputString)
        If Mid(inputString, i, 1) = charToCount Then
            count = count + 1
        End If
    Next i
    
    Fn_Count_Char_Occurrences = count
    
End Function

'Fonction de tri rapide (QuickSort) pour trier un tableau
Sub Fn_Quick_Sort(arr As Variant, ByVal first As Long, ByVal last As Long) '2024-09-05 @ 05:09
    
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
        If first < j Then Fn_Quick_Sort arr, first, j
        If i < last Then Fn_Quick_Sort arr, i, last
    End If
    
End Sub

Function Fn_Numero_Semaine_Selon_AnneeFinancière(DateDonnee As Date) As Long
    
    Dim DebutAnneeFinanciere As Date
    DebutAnneeFinanciere = wshAdmin.Range("AnneeDe")
    
    'Trouver le jour de la semaine du début de l'année financière (1 = dimanche, 2 = lundi, etc.)
    Dim JourSemaineDebut As Long
    JourSemaineDebut = Weekday(DebutAnneeFinanciere, vbMonday)
    
    ' Ajuster le début de l'année financière au lundi précédent si ce n'est pas un lundi
    If JourSemaineDebut > 1 Then
        DebutAnneeFinanciere = DebutAnneeFinanciere - (JourSemaineDebut - 1)
    End If
    
    ' Calculer le nombre de jours entre la date donnée et le début ajusté de l'année financière
    Dim NbJours As Long
    NbJours = DateDonnee - DebutAnneeFinanciere
    
    ' Calculer le numéro de la semaine (diviser par 7 et arrondir)
    Dim Semaine As Integer
    Semaine = Int(NbJours / 7) + 1
    
    ' Retourner le numéro de la semaine
    Fn_Numero_Semaine_Selon_AnneeFinancière = Semaine
    
End Function

Function Fn_Valider_Portion_Heures(valeur As Currency) As Boolean

    'Tableau des valeurs permises : dixièmes d'heures et quarts d'heure
    Dim valeursPermises As Variant
    valeursPermises = Array(0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 0.25, 0.75)
    
    Dim i As Integer
    Fn_Valider_Portion_Heures = False 'Initialisation à Faux
    
    'Parcourir les valeurs permises
    Dim fraction As Double
    fraction = valeur - Int(valeur)
    
    For i = LBound(valeursPermises) To UBound(valeursPermises)
        If Round(fraction, 2) = valeursPermises(i) Then
            Fn_Valider_Portion_Heures = True 'La fraction est valide
            Exit Function
        End If
    Next i
    
End Function

Function Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arrière(d As Date) As Date

    'Cette fonction calcule le premier jour du trimestre pour une date de fin de trimestre (TPS/TVQ)
    Dim dateTroisMoisAvant As Date
    
    'Reculer de trois mois à partir de la date saisie
    dateTroisMoisAvant = DateAdd("m", -2, d)
    
    'Fixer le jour au PREMIER du mois obtenu
    Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arrière = DateSerial(year(dateTroisMoisAvant), month(dateTroisMoisAvant), 1)
    
End Function

Function Fn_Obtenir_Date_Lundi(d As Date)

    Fn_Obtenir_Date_Lundi = d - (Weekday(d, vbMonday) - 1)

End Function

Function Fn_Nettoyer_Fin_Chaine(s As String) '2024-11-07 @ 16:57

    Fn_Nettoyer_Fin_Chaine = s
    
    'Supprimer les retours à la ligne, les sauts de ligne et les espaces inutiles
    Fn_Nettoyer_Fin_Chaine = Trim(Replace(Replace(Replace(s, vbCrLf, ""), vbCr, ""), vbLf, ""))

End Function

'Fonction pour vérifier si un fichier ou un dossier existe
Private Function Fn_Chemin_Existe(ByVal chemin As String) As Boolean

    On Error Resume Next
    Fn_Chemin_Existe = (Dir(chemin) <> "")
    On Error GoTo 0
    
End Function

Function ObtenirNoGlIndicateur(ByVal indic As Variant) As String

    'Plage où sont situés les liens (indicateur/no de GL)
    Dim plage As Range
    Set plage = wshAdmin.Range("D44:F60")
    
    'Parcourir chaque cellule dans la première colonne de la plage
    Dim cellule As Range
    For Each cellule In plage.Columns(1).Cells
        If cellule.value = indic Then
            'Retourner la valeur de la troisième colonne pour la ligne correspondante
            ObtenirNoGlIndicateur = cellule.offset(0, 1).value
            Exit Function
        End If
    Next cellule
    
    'Si la valeur n'est pas trouvée
    ObtenirNoGlIndicateur = "Non trouvé"

End Function

Function Fn_Obtenir_Paiements_Facture(invNo As String, dateLimite As Date) As Currency

    Dim ws As Worksheet: Set ws = wshENC_Détails
    
    Dim premièreCellule As Range, celluleTrouvée As Range
    Dim ligne As Long
    
    'Rechercher la première cellule avec le numéro de facture
    Set premièreCellule = ws.Columns(2).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    Fn_Obtenir_Paiements_Facture = 0
    
    If Not premièreCellule Is Nothing Then
        'Démarrer la recherche à partir de la première cellule trouvée
        Set celluleTrouvée = premièreCellule
        Do
            'Obtenir la ligne correspondante
            ligne = celluleTrouvée.row
            'Vérifier la date dans la colonne appropriée
            If IsDate(ws.Cells(ligne, 4).value) And ws.Cells(ligne, 4).value <= dateLimite Then
                'Additionner le montant du paiement
                Fn_Obtenir_Paiements_Facture = Fn_Obtenir_Paiements_Facture + ws.Cells(ligne, 5).value
            End If
            'Rechercher la prochaine occurrence
            Set celluleTrouvée = ws.Columns(2).FindNext(celluleTrouvée)
        Loop While Not celluleTrouvée Is Nothing And celluleTrouvée.Address <> premièreCellule.Address
    End If

End Function

Function Fn_Obtenir_Régularisations_Facture(invNo As String, dateLimite As Date) As Currency

    Dim ws As Worksheet: Set ws = wshCC_Régularisations
    
    Dim premièreCellule As Range, celluleTrouvée As Range
    Dim ligne As Long
    
    'Rechercher la première cellule avec le numéro de facture
    Set premièreCellule = ws.Columns(fREGULInvNo).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    Fn_Obtenir_Régularisations_Facture = 0
    
    If Not premièreCellule Is Nothing Then
        'Démarrer la recherche à partir de la première cellule trouvée
        Set celluleTrouvée = premièreCellule
        Do
            'Obtenir la ligne correspondante
            ligne = celluleTrouvée.row
            'Vérifier la date dans la colonne appropriée
            If IsDate(ws.Cells(ligne, fREGULDate).value) And ws.Cells(ligne, fREGULDate).value <= dateLimite Then
                'Additionner les cellules pertinentes
                Fn_Obtenir_Régularisations_Facture = Fn_Obtenir_Régularisations_Facture + _
                                                        ws.Cells(ligne, fREGULHono).value + _
                                                        ws.Cells(ligne, fREGULFrais).value + _
                                                        ws.Cells(ligne, fREGULTPS).value + _
                                                        ws.Cells(ligne, fREGULTVQ).value
            End If
            'Rechercher la prochaine occurrence
            Set celluleTrouvée = ws.Columns(fREGULInvNo).FindNext(celluleTrouvée)
        Loop While Not celluleTrouvée Is Nothing And celluleTrouvée.Address <> premièreCellule.Address
    End If

End Function

Function ExtraireSecondes(chaine As String) As Double
    
    Dim pos As Integer
    Dim secondes As String
    
    chaine = Replace(chaine, ".", ",")
    chaine = Replace(chaine, "'", "")
    
    If InStr(chaine, " = ") > 0 Then
        chaine = Right(chaine, Len(chaine) - InStr(chaine, " = ") - 2)
    End If
    
    'Trouve la position de " secondes"
    pos = InStr(chaine, " secondes")
    If pos > 0 Then
        secondes = Left(chaine, pos - 1)
        ExtraireSecondes = CStr(secondes)
    Else
        'Si " secondes" n'est pas trouvé, retourne une chaîne vide
        ExtraireSecondes = "0"
    End If
    
End Function

'Fonction pour centraliser tous les messages de l'application - 2024-12-29 @ 07:37
Function AppMsgBox(message As String _
                 , Optional boutons As VbMsgBoxStyle = vbOKOnly _
                 , Optional titre As String = "") As VbMsgBoxResult
                 
    AppMsgBox = MsgBox(message, boutons, titre)
                 
End Function

Sub test_AppMsgBox()

    Dim r As VbMsgBoxResult
    r = AppMsgBox("Voulez-vous continuer ?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirmation avant de continuer")

    Debug.Print "#090 - " & r
    
End Sub

Function ExtraireNomFichier(path As String) As String

    Dim parts() As String
    parts = Split(path, "\")
    
    ExtraireNomFichier = parts(UBound(parts, 1))

End Function

Function ConvertirEnNumerique(rng As Range)

    Dim cell As Range
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CCur(cell.value)
        Else
            cell.value = CCur(Replace(cell.value, " ", ""))
        End If
    Next cell
    
End Function

Function TrouverLigneFacture(ws As Worksheet, numFacture As String) As Long

    Dim rng As Range
    Set rng = ws.Range("A:A").Find(What:=numFacture, LookAt:=xlWhole)

    If Not rng Is Nothing Then
        TrouverLigneFacture = rng.row
    Else
        TrouverLigneFacture = -1 ' Retourne -1 si la facture n'est pas trouvée
    End If
    
End Function

Function RechercherLignesTableau(ws As Worksheet, noEntrée As Long) As Variant

    'Définition de la feuille contenant les données
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    Dim plageRecherche As Range
    Set plageRecherche = ws.Range("A2:A" & lastUsedRow)

    'Recherche des lignes correspondantes
    Dim plageResultat As Range
    Dim cell As Range
    For Each cell In plageRecherche
        If cell.value = noEntrée Then
            If plageResultat Is Nothing Then
                Set plageResultat = cell.EntireRow
            Else
                Set plageResultat = Union(plageResultat, cell.EntireRow)
            End If
        End If
    Next cell

    ' Vérifier si des résultats ont été trouvés
    If plageResultat Is Nothing Then
        RechercherLignesTableau = False ' Aucun résultat trouvé
        Exit Function
    End If

    'Convertir la plage trouvée en tableau VBA
    Dim nbLignes As Integer, nbColonnes As Integer
    Dim data As Variant
    nbLignes = plageResultat.Rows.count
    nbColonnes = plageResultat.Columns.count
    data = plageResultat.value ' Stocker la plage dans un tableau Variant

    ' Retourner le tableau des résultats
    RechercherLignesTableau = data
    
End Function

Function EstLigneSelectionnee(ByVal lb As Object) As Boolean

    Dim i As Integer
    EstLigneSelectionnee = False 'Par défaut, aucune ligne n'est sélectionnée

    'Vérifier toutes les lignes du ListBox
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            EstLigneSelectionnee = True 'Si une ligne est sélectionnée
            Exit Function 'Quitter dès qu'une sélection est trouvée
        End If
    Next i
    
End Function

