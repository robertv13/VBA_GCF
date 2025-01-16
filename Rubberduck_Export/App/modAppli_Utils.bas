Attribute VB_Name = "modAppli_Utils"
Option Explicit

'Variables globales pour le module
Dim verificationIntegriteOK As Boolean
Dim soldeComptesClients As Currency

'@Folder("")

Public Sub ConvertRangeBooleanToText(rng As Range)

    Dim cell As Range
    For Each cell In rng
        Select Case cell.Value
            Case 0, "False" 'False
                cell.Value = "FAUX"
            Case -1, "True" 'True
                cell.Value = "VRAI"
            Case "VRAI", "FAUX"
                
            Case Else
                MsgBox cell.Value & " est une valeur INVALIDE pour la cellule " & cell.Address & " de la feuille TEC_Local"
        End Select
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
End Sub

Public Sub VérifierIntégrité() '2024-11-20 @ 06:55

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:Integrity_Verification", 0)

    Application.ScreenUpdating = False
    
    'Variable pour déterminer à la fin s'il y a des erreurs...
    verificationIntegriteOK = True
    
    Call Erase_And_Create_Worksheet("X_Analyse_Intégrité")
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    wsOutput.Unprotect
    wsOutput.Range("A1").Value = "Feuille"
    wsOutput.Range("B1").Value = "Message"
    wsOutput.Range("C1").Value = "TimeStamp"
    wsOutput.Columns("C").NumberFormat = wshAdmin.Range("B1").Value & " hh:mm:ss"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"))

    Application.ScreenUpdating = True

    'Data starts at row 2
    Dim r As Long: r = 2
    Call AddMessageToWorkSheet(wsOutput, r, 1, "Répertoire utilisé")
    Call AddMessageToWorkSheet(wsOutput, r, 2, wshAdmin.Range("FolderSharedData").Value & DATA_PATH)
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1

    'Fichier utilisé
    Dim masterFileName As String
    masterFileName = "GCF_BD_MASTER.xlsx"
    Call AddMessageToWorkSheet(wsOutput, r, 1, "Fichier utilisé")
    Call AddMessageToWorkSheet(wsOutput, r, 2, masterFileName)
    r = r + 1
    
    'Date dernière modification du fichier MAÎTRE
    Dim fullFileName As String
    fullFileName = wshAdmin.Range("FolderSharedData").Value & DATA_PATH & Application.PathSeparator & masterFileName
    Dim ddm As Date
    Dim j As Long, h As Long, m As Long, s As Long
    Call Get_Date_Derniere_Modification(fullFileName, ddm, j, h, m, s)
    Call AddMessageToWorkSheet(wsOutput, r, 1, "Date dern. modification")
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = Format$(ddm, wshAdmin.Range("B1").Value & " hh:mm:ss") & _
            " soit " & j & " jours, " & h & " heures, " & m & " minutes et " & s & " secondes"
    rng.Characters(1, 19).Font.COLOR = vbRed
    rng.Characters(1, 19).Font.Bold = True

    r = r + 2
    
    Dim readRows As Long
    
    'dnrPlanComptable ----------------------------------------------------- Plan Comptable
    Call AddMessageToWorkSheet(wsOutput, r, 1, "Plan Comptable")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    
    Call checkPlanComptable(r, readRows)
    
    'wshBD_Clients --------------------------------------------------------------- Clients
    Call AddMessageToWorkSheet(wsOutput, r, 1, "BD_Clients")
    
    Call Client_List_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "La feuille a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkClients(r, readRows)
    
    'wshBD_Fournisseurs ----------------------------------------------------- Fournisseurs
    Call AddMessageToWorkSheet(wsOutput, r, 1, "BD_Fournisseurs")
    
    Call Fournisseur_List_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "La feuille a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkFournisseurs(r, readRows)
    
    'wshDEB_Récurrent ------------------------------------------------------ DEB_Récurrent
    Call AddMessageToWorkSheet(wsOutput, r, 1, "DEB_Récurrent")
    
    Call DEB_Récurrent_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "DEB_Récurrent a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkDEB_Récurrent(r, readRows)
    
    'wshDEB_Trans -------------------------------------------------------------- DEB_Trans
    Call AddMessageToWorkSheet(wsOutput, r, 1, "DEB_Trans")
    
    Call DEB_Trans_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "DEB_Trans a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkDEB_Trans(r, readRows)
    
    'wshFAC_Entête ------------------------------------------------------------ FAC_Entête
    Call AddMessageToWorkSheet(wsOutput, r, 1, "FAC_Entête")
    
    Call FAC_Entête_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "FAC_Entête a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkFAC_Entête(r, readRows)
    
    'wshFAC_Détails ---------------------------------------------------------- FAC_Détails
    Call AddMessageToWorkSheet(wsOutput, r, 1, "FAC_Détails")
    
    Call FAC_Détails_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "FAC_Détails a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkFAC_Détails(r, readRows)
    
    'wshFAC_Comptes_Clients ------------------------------------------ FAC_Comptes_Clients
    Call AddMessageToWorkSheet(wsOutput, r, 1, "FAC_Comptes_Clients")
    
    Call FAC_Comptes_Clients_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "FAC_Comptes_Clients a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkFAC_Comptes_Clients(r, readRows)
    
    'wshFAC_Sommaire_Taux ---------------------------------------------- FAC_Sommaire_Taux
    Call AddMessageToWorkSheet(wsOutput, r, 1, "FAC_Sommaire_Taux")
    
    Call FAC_Sommaire_Taux_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "FAC_Sommaire_Taux a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkFAC_Sommaire_Taux(r, readRows)
    
    'wshENC_Entête ------------------------------------------------------------ ENC_Entête
    Call AddMessageToWorkSheet(wsOutput, r, 1, "ENC_Entête")
    
    Call ENC_Entête_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "ENC_Entête a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkENC_Entête(r, readRows)
    
    'wshENC_Détails ---------------------------------------------------------- ENC_Détails
    Call AddMessageToWorkSheet(wsOutput, r, 1, "ENC_Détails")
    
    Call ENC_Détails_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "ENC_Détails a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkENC_Détails(r, readRows)
    
    'wshCC_Régularisations ---------------------------------------------------------- CC_Régularisations
    Call AddMessageToWorkSheet(wsOutput, r, 1, "CC_Régularisations")
    
    Call CC_Régularisations_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "CC_Régularisations a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkCC_Régularisations(r, readRows)
    
    'wshFAC_Projets_Entête -------------------------------------------- FAC_Projets_Entête
    Call AddMessageToWorkSheet(wsOutput, r, 1, "FAC_Projets_Entête")
    
    Call FAC_Projets_Entête_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "FAC_Projets_Entête a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkFAC_Projets_Entête(r, readRows)
    
    'wshFAC_Projets_Détails ------------------------------------------ FAC_Projets_Détails
    Call AddMessageToWorkSheet(wsOutput, r, 1, "FAC_Projets_Détails")
    
    Call FAC_Projets_Détails_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "FAC_Projets_Détails a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkFAC_Projets_Détails(r, readRows)
    
    'wshGL_Trans ---------------------------------------------------------------- GL_Trans
    Call AddMessageToWorkSheet(wsOutput, r, 1, "GL_Trans")
    
    Call GL_Trans_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "GL_Trans a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkGL_Trans(r, readRows)
    
    'wshGL_EJ_Recurrente ------------------------------------------------ GL_EJ_Recurrente
    Call AddMessageToWorkSheet(wsOutput, r, 1, "GL_EJ_Recurrente")
    
    Call GL_EJ_Recurrente_Import_All
    Call AddMessageToWorkSheet(wsOutput, r, 2, "GL_EJ_Recurrente a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkGL_EJ_Recurrente(r, readRows)
    
    'wshTEC_TdB_Data -------------------------------------------------------- TEC_TdB_Data
    Call AddMessageToWorkSheet(wsOutput, r, 1, "TEC_TdB_Data")
    
    Call TEC_Import_All
    Call modTEC_TDB.ActualiserTEC_TDB
    Call AddMessageToWorkSheet(wsOutput, r, 2, "TEC_TdB_Data a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkTEC_TdB_Data(r, readRows)
    
    'wshTEC_Local -------------------------------------------------------------- TEC_Local
    Call AddMessageToWorkSheet(wsOutput, r, 1, "TEC_Local")
    Call AddMessageToWorkSheet(wsOutput, r, 2, "TEC_Local a été importée du fichier BD_MASTER.xlsx")
    Call AddMessageToWorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    r = r + 1
    
    Call checkTEC(r, readRows)
    
    'Adjust the Output Worksheet
    With wsOutput.Range("A2:C" & r).Font
        .Name = "Courier New"
        .size = 10
    End With
    
    wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
   'Result print setup - 2024-07-20 @ 14:31
    Dim lastUsedRow As Long
    lastUsedRow = r
    
    'Un peu de couleur
    Set rng = wsOutput.Range("A" & r)
    rng.Value = "**** " & Format$(readRows, "###,##0") & _
                    " lignes analysées dans l'ensemble des tables - " & _
                    Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss") & " ***"
    rng.Characters(6, 6).Font.COLOR = vbRed
    rng.Characters(6, 6).Font.Bold = True
    
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "Vérification d'intégrité des tables"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    If verificationIntegriteOK = True Then
        MsgBox "La vérification d'intégrité est terminé SANS PROBLÈME" & vbNewLine & vbNewLine & "Voir la feuille 'X_Analyse_Intégrité'", vbInformation
    Else
        MsgBox "La vérification a détecté AU MOINS UN PROBLÈME" & vbNewLine & vbNewLine & "Voir la feuille 'X_Analyse_Intégrité'", vbInformation
    End If
    
    ThisWorkbook.Worksheets("X_Analyse_Intégrité").Activate
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set rng = Nothing
    Set rngToPrint = Nothing
    Set wsOutput = Nothing
    
    Call Log_Record("modAppli:Integrity_Verification", startTime)

End Sub

Sub Simple_Print_Setup(ws As Worksheet, rng As Range, header1 As String, _
                       header2 As String, titleRows As String, Optional Orient As String = "L")
    
    On Error GoTo CleanUp
    
    Application.PrintCommunication = False
    
    With ws.PageSetup
        .PrintArea = rng.Address
        .PrintTitleRows = titleRows
        .PrintTitleColumns = ""
        
        .CenterHeader = "&""-,Gras""&12&K0070C0" & header1 & Chr(10) & "&11" & header2
        
        .LeftFooter = "&8&D - &T"
        .CenterFooter = "&8&KFF0000&A"
        .RightFooter = "&""Segoe UI,Normal""&8Page &P of &N"
        
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1.5)
        
        .BottomMargin = Application.CentimetersToPoints(1)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        
        .CenterHorizontally = True
        
        If Orient = "L" Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .PaperSize = xlPaperLetter
        .FitToPagesWide = 1
        .FitToPagesTall = 10
    End With
    
CleanUp:
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
    
End Sub

Public Sub ProtectCells(rng As Range)

    'Lock the range
    rng.Locked = True
    
    'Protect the worksheet with no restrictions
    With rng.Parent
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlNoRestrictions
    End With

End Sub

Public Sub UnprotectCells(rng As Range)

    'Unlcok the range
    rng.Locked = False
    
    'Protect the worksheet with no restrictions
    With rng.Parent
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlNoRestrictions
    End With

End Sub

Public Sub Tx_2D_Array_2_Range(ByRef arr As Variant, _
                               ByVal rngTo As Range, _
                               Optional ByVal clearExistingData As Boolean = True, _
                               Optional ByVal HeaderSize As Long = 1)
                        
    'Si requis, on efface le contenu de rngTo avant
    If clearExistingData = True Then
        rngTo.CurrentRegion.offset(HeaderSize).ClearContents
    End If
    
    'En fonction des dimensions du tableau (arr)
    Dim r As Long, c As Long
    r = UBound(arr, 1) - LBound(arr, 1) + HeaderSize
    c = UBound(arr, 2) - LBound(arr, 2) + HeaderSize
    rngTo.Resize(r, c).Value = arr
    
End Sub

Sub Tx_Range_2_2D_Array(ByVal rng As Range, ByRef arr As Variant, Optional ByVal headerRows As Long = 1)

    'La plage est-elle valide ?
    If rng Is Nothing Then
        MsgBox "La plage est invalide ou non définie.", vbExclamation, , "modAppli_Utils:Tx_Range_2_2D_Array"
        Exit Sub
    End If
    
    'Calculer la taille de la plage des données pour ensuite ignorer les en-têtes
    Dim numRows As Long, numCols As Long
'    startRow = rng.row + headerRows
    numRows = rng.Rows.count - headerRows
    numCols = rng.Columns.count
    
    'La plage contient-elle des données ?
    If numRows <= 0 Or numCols <= 0 Then
        MsgBox "Aucune donnée à copier dans le tableau.", vbExclamation, "modAppli_Utils:Tx_Range_2_2D_Array"
        Exit Sub
    End If
    
    'Définir la taille de la plage qui contient les données, en fonction de numRows & numCols
    On Error Resume Next
    Dim rngData As Range
    Set rngData = rng.Resize(numRows, numCols).offset(headerRows, 0)
    On Error GoTo 0
    
    'Copier les données du Rage vers le tableau (Array)
    If Not rngData Is Nothing Then
        arr = rngData.Value
    Else
        MsgBox "Erreur lors de la création de la plage de données.", vbExclamation, "modAppli_Utils:Tx_Range_2_2D_Array"
    End If
    
    'Libérer la mémoire
    Set rngData = Nothing
    
End Sub

Sub CreateOrReplaceWorksheet(wsName As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli_Utils:CreateOrReplaceWorksheet", 0)
    
    'Check if the worksheet exists
    Dim wsExists As Boolean
    wsExists = NomFeuilleExiste(wsName)
    
    'If the worksheet exists, delete it
    If wsExists Then
        Application.DisplayAlerts = False
        Worksheets(wsName).Delete
        Application.DisplayAlerts = True
    End If
    
    'Add the new worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=wshMenu)
    ws.Name = wsName

    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modAppli_Utils:CreateOrReplaceWorksheet", startTime)

End Sub
Private Sub checkPlanComptable(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkPlanComptable", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'dnrPlanComptable_All
    Dim arr As Variant
    Dim nbCol As Long
    nbCol = 4
    arr = Fn_Get_Plan_Comptable(nbCol) 'Returns array with 4 columns (Code, Description)
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(UBound(arr, 1), "###,##0") & _
        " comptes et " & Format$(nbCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de 'dnr_PlanComptable_All'")
    r = r + 1
    
    If UBound(arr, 1) < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim dict_code_GL As New Dictionary
    Dim dict_descr_GL As New Dictionary
    
    Dim i As Long, codeGL As String, descrGL As String
'    Dim GL_ID As Long
    Dim typeGL As String
    Dim cas_doublon_descr As Long, cas_doublon_code As Long, cas_type As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        codeGL = arr(i, 1)
        descrGL = arr(i, 2)
        If dict_descr_GL.Exists(descrGL) = False Then
            dict_descr_GL.Add descrGL, codeGL
        Else
            Call AddMessageToWorkSheet(wsOutput, r, 2, "La description '" & descrGL & "' est un doublon pour le code de G/L '" & codeGL & "'")
            r = r + 1
            cas_doublon_descr = cas_doublon_descr + 1
        End If
        
        If dict_code_GL.Exists(codeGL) = False Then
            dict_code_GL.Add codeGL, descrGL
        Else
            Call AddMessageToWorkSheet(wsOutput, r, 2, "Le code de G/L '" & codeGL & "' est un doublon pour la description '" & descrGL & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
        
'        GL_ID = arr(i, 3)
        typeGL = arr(i, 4)
        If InStr("Actifs^Passifs^Équité^Revenus^Dépenses^", typeGL) = 0 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "Le type de compte '" & typeGL & "' est INVALIDE pour le code de G/L '" & codeGL & "'")
            r = r + 1
            cas_type = cas_type + 1
        End If
        
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " comptes ont été analysés!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_descr = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucun doublon de description")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_doublon_descr & " cas de doublons pour les descriptions")
        r = r + 1
    End If
    
    If cas_doublon_code = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucun doublon de code de G/L")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_doublon_code & " cas de doublons pour les codes de G/L")
        r = r + 1
    End If
    
    If cas_type = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucun type de G/L invalide")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_type & " cas de types de G/L invalides")
        r = r + 1
    End If
    r = r + 1
    
    'Cas problème dans cette vérification ?
    If cas_doublon_descr <> 0 Or cas_doublon_descr <> 0 Or cas_type <> 0 Then
        verificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkPlanComptable", startTime)

End Sub

Private Sub checkClients(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkClients", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'Fichier maître des Clients
    Dim ws As Worksheet: Set ws = wshBD_Clients
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(ws.Range("A1").CurrentRegion.Rows.count, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshBD_Clients'")
    r = r + 1
    
    Dim arr As Variant
    arr = wshBD_Clients.Range("A1").CurrentRegion.Value
    If UBound(arr, 1) < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim dict_code_client As New Dictionary
    Dim dict_nom_client As New Dictionary
    
    Dim i As Long, code As String, nom As String, nomClientSysteme As String
    Dim cas_doublon_nom As Long
    Dim cas_doublon_code As Long
    Dim cas_courriel_invalide As Long
    Dim ligneNonVides As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        If Not Trim(arr(i, 2)) = "" Then
        ligneNonVides = ligneNonVides + 1
            nom = arr(i, 1)
            code = arr(i, 2)
            nomClientSysteme = arr(i, 3)
            
            'Doublon sur le nom ?
            If dict_nom_client.Exists(nom) = False Then
                dict_nom_client.Add nom, code
            Else
                Call AddMessageToWorkSheet(wsOutput, r, 2, "À la ligne " & i & ", le nom '" & nom & "' est un doublon pour le code '" & code & "'")
                r = r + 1
                cas_doublon_nom = cas_doublon_nom + 1
            End If
            
            'Doublon sur le code de client ?
            If dict_code_client.Exists(code) = False Then
                dict_code_client.Add code, nom
            Else
                Call AddMessageToWorkSheet(wsOutput, r, 2, "À la ligne " & i & ", le code '" & code & "' est un doublon pour le client '" & nom & "'")
                r = r + 1
                cas_doublon_code = cas_doublon_code + 1
            End If
            
            If Trim(arr(i, 6)) <> "" Then
                If Fn_ValiderCourriel(arr(i, 6)) = False Then
                    Call AddMessageToWorkSheet(wsOutput, r, 2, "À la ligne " & i & ", le courriel '" & arr(i, 6) & "' est INVALIDE pour le code '" & code & "'")
                    r = r + 1
                    cas_courriel_invalide = cas_courriel_invalide + 1
                End If
            End If
        End If
    Next i
    
    'Toutes les lignes sont-elles non-vides ?
    If UBound(arr, 1) - 1 <> ligneNonVides Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La feuille comporte un ou des ligne(s) vide(s)")
        r = r + 1
    End If
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Un total de " & Format$(UBound(arr, 1) - 1, "##,##0") & " clients ont été analysés!"
    rng.Characters(13, 5).Font.COLOR = vbRed
    rng.Characters(13, 5).Font.Bold = True

    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucun doublon de nom")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        r = r + 1
    End If
    
    If cas_doublon_code = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucun doublon de code")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        r = r + 1
    End If
    
    If cas_courriel_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Toutes les adresses courriel sont valides")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_courriel_invalide & " cas de courriels INVALIDES")
        r = r + 1
    End If
    
    r = r + 1
    
    'Cas problème dans cette vérification ?
    If cas_doublon_nom <> 0 Or _
        cas_doublon_code <> 0 Or _
        cas_courriel_invalide <> 0 Or _
        UBound(arr, 1) - 1 <> ligneNonVides Then
            verificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set dict_code_client = Nothing
    Set dict_nom_client = Nothing
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkClients", startTime)

End Sub

Private Sub checkFournisseurs(ByRef r As Long, ByRef readRows As Long)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkFournisseurs", 0)

    Application.ScreenUpdating = False

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshBD_fournisseurs
    Dim ws As Worksheet: Set ws = wshBD_Fournisseurs
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(ws.Range("A1").CurrentRegion.Rows.count, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshBD_Fournisseurs'")
    r = r + 1
    
    Dim arr As Variant
    arr = wshBD_Fournisseurs.Range("A1").CurrentRegion.Value
    If UBound(arr, 1) < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim dict_code_fournisseur As New Dictionary
    Dim dict_nom_fournisseur As New Dictionary
    
    Dim i As Long, code As String, nom As String
    Dim cas_doublon_nom As Long
    Dim cas_doublon_code As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        nom = arr(i, 1)
        code = arr(i, 2)
        If dict_nom_fournisseur.Exists(nom) = False Then
            dict_nom_fournisseur.Add nom, code
        Else
            Call AddMessageToWorkSheet(wsOutput, r, 2, "Le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        If dict_code_fournisseur.Exists(code) = False Then
            dict_code_fournisseur.Add code, nom
        Else
            Call AddMessageToWorkSheet(wsOutput, r, 2, "Le code '" & code & "' est un doublon pour le nom '" & nom & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
    Next i
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Un total de " & Format$(UBound(arr, 1) - 1, "#,##0") & " fournisseurs ont été analysés!"
    rng.Characters(13, 3).Font.COLOR = vbRed
    rng.Characters(13, 3).Font.Bold = True

    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucun doublon de nom")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        r = r + 1
    End If
    If cas_doublon_code = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucun doublon de code")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        r = r + 1
    End If
    r = r + 1
    
    'Cas problème dans cette vérification ?
    If cas_doublon_nom <> 0 Or cas_doublon_code <> 0 Then
        verificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkFournisseurs", startTime)

End Sub

Private Sub checkDEB_Récurrent(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkDEB_Récurrent", 0)

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshDEB_Récurrent
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshDEB_Récurrent
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= 2 - HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshDEB_Récurrent'")
    r = r + 1
    
    'On a besoin du plan comptable pour valider les données
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wshAdmin.Range("dnrPlanComptable_All")
    On Error GoTo 0

    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        r = r + 1
        Exit Sub
    End If
    
    Dim strGL As String
    Dim ligne As Range
    For Each ligne In planComptable.Rows
        strGL = strGL & "^C:" & Trim(ligne.Cells(1, 2).Value) & "^D:" & Trim(ligne.Cells(1, 1).Value) & " | "
    Next ligne
    
    'Copie les données vers un tableau
    Dim rng As Range
    Set rng = ws.Range("A1:N" & lastUsedRow)
    Dim arr() As Variant
    Dim headerRows As Long
    headerRows = 1
    Call Tx_Range_2_2D_Array(rng, arr, 1)
    
    'On analyse chacune des lignes du tableau
    Dim i As Long, p As Long
    Dim GL As String, descGL As String
    
    Dim isDebRécurrentValid As Boolean
    isDebRécurrentValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsNumeric(arr(i, 1)) = False Or arr(i, 1) <> Int(arr(i, 1)) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la valeur du numéro de déboursé '" & arr(i, 1) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        
        If IsDate(arr(i, 2)) = False Or arr(i, 2) <> Int(arr(i, 2)) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la date '" & arr(i, 2) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        
        p = InStr(strGL, "^C:" & arr(i, 6))
        If p = 0 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le poste de G/L '" & arr(i, 6) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        If p > 0 Then
            GL = Mid(strGL, p + 3)
            descGL = Mid(GL, InStr(GL, "^D:") + 3, InStr(GL, " | ") - 8)
            If descGL <> Trim(arr(i, 7)) Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la description du G/L '" & arr(i, 7) & "' est INVALIDE")
                r = r + 1
                isDebRécurrentValid = False
            End If
        End If
        
        'Total
        If IsNumeric(arr(i, 9)) = False Or arr(i, 9) * 100 <> Int(arr(i, 9) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant total '" & arr(i, 9) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        'TPS
        If IsNumeric(arr(i, 10)) = False Or arr(i, 10) * 100 <> Int(arr(i, 10) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant TPS '" & arr(i, 10) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        'TVQ
        If IsNumeric(arr(i, 11)) = False Or arr(i, 11) * 100 <> Int(arr(i, 11) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant TVQ '" & arr(i, 11) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        'Intrant TPS
        If IsNumeric(arr(i, 12)) = False Or arr(i, 12) * 100 <> Int(arr(i, 12) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant d'intrant pour la TPS '" & arr(i, 12) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        'Intrant TVQ
        If IsNumeric(arr(i, 13)) = False Or arr(i, 13) * 100 <> Int(arr(i, 13) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant d'intrant pour la TVQ '" & arr(i, 13) & "' est INVALIDE")
            r = r + 1
            isDebRécurrentValid = False
        End If
        '$ dépense
        readRows = readRows + 1
    Next i

    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Cas problème dans cette vérification ?
    If isDebRécurrentValid = False Then
        verificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set ligne = Nothing
    Set planComptable = Nothing
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing

    Application.ScreenUpdating = True

    Call Log_Record("modAppli:checkDEB_Récurrent", startTime)

End Sub

Private Sub checkDEB_Trans(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkDEB_Trans", 0)

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshDEB_Trans
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshDEB_Trans
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= 2 - HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshDEB_Trans'")
    r = r + 1
    
    'On a besoin du plan comptable pour valider les données
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wshAdmin.Range("dnrPlanComptable_All")
    On Error GoTo 0

    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        r = r + 1
        Exit Sub
    End If
    
    Dim strGL As String
    Dim ligne As Range
    For Each ligne In planComptable.Rows
        strGL = strGL & "^C:" & Trim(ligne.Cells(1, 2).Value) & "^D:" & Trim(ligne.Cells(1, 1).Value) & " | "
    Next ligne
    
    'Copie les données vers un tableau
    Dim rng As Range
    Set rng = ws.Range("A1:R" & lastUsedRow)
    Dim arr() As Variant
    Dim headerRows As Long
    headerRows = 1
    Call Tx_Range_2_2D_Array(rng, arr, 1)
    
    'On analyse chacune des lignes du tableau
    Dim i As Long, p As Long
    Dim GL As String, descGL As String
    
    Dim isDebTransValid As Boolean
    isDebTransValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsNumeric(arr(i, 1)) = False Or arr(i, 1) <> Int(arr(i, 1)) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la valeur du numéro de déboursé '" & arr(i, 1) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        
        If IsDate(arr(i, 2)) = False Or arr(i, 2) <> Int(arr(i, 2)) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la date '" & arr(i, 2) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        
        p = InStr(strGL, "^C:" & arr(i, 8))
        If p = 0 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le poste de G/L '" & arr(i, 8) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        If p > 0 Then
            GL = Mid(strGL, p + 3)
            descGL = Mid(GL, InStr(GL, "^D:") + 3, InStr(GL, " | ") - 8)
            If descGL <> Trim(arr(i, 9)) Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la description du G/L '" & arr(i, 8) & "' est INVALIDE")
                r = r + 1
                isDebTransValid = False
            End If
        End If
        
        'Total
        If IsNumeric(arr(i, 11)) = False Or arr(i, 11) * 100 <> Int(arr(i, 11) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant total '" & arr(i, 11) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        'TPS
        If IsNumeric(arr(i, 12)) = False Or arr(i, 12) * 100 <> Int(arr(i, 12) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant TPS '" & arr(i, 12) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        'TVQ
        If IsNumeric(arr(i, 13)) = False Or arr(i, 13) * 100 <> Int(arr(i, 13) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant TVQ '" & arr(i, 13) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        'Intrant TPS
        If IsNumeric(arr(i, 14)) = False Or arr(i, 14) * 100 <> Int(arr(i, 14) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant d'intrant pour la TPS '" & arr(i, 14) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        'Intrant TVQ
        If IsNumeric(arr(i, 15)) = False Or arr(i, 15) * 100 <> Int(arr(i, 15) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant d'intrant pour la TVQ '" & arr(i, 15) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        '$ dépense
        If IsNumeric(arr(i, 16)) = False Or arr(i, 16) * 100 <> Int(arr(i, 16) * 100) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le montant de la dépense '" & arr(i, 16) & "' est INVALIDE")
            r = r + 1
            isDebTransValid = False
        End If
        readRows = readRows + 1
    Next i

    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Cas problème dans cette vérification ?
    If isDebTransValid = False Then
        verificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set ligne = Nothing
    Set planComptable = Nothing
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing

    Application.ScreenUpdating = True

    Call Log_Record("modAppli:checkDEB_Trans", startTime)

End Sub

Private Sub checkENC_Détails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkENC_Détails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshENC_Détails
    Dim ws As Worksheet: Set ws = wshENC_Détails
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRowDetails As Long
    lastUsedRowDetails = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRowDetails <= 2 - HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRowDetails, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    'ENC_Entête Worksheet
    Dim wsEntete As Worksheet: Set wsEntete = wshENC_Entête
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsEntete.Cells(wsEntete.Rows.count, 1).End(xlUp).row
    Dim rngEntete As Range: Set rngEntete = wsEntete.Range("A2:A" & lastUsedRowEntete)
    Dim strPmtNo As String
    Dim i As Long
    For i = 2 To lastUsedRowEntete
        strPmtNo = strPmtNo & CLng(wsEntete.Cells(i, fEncEPayID).Value) & "|"
    Next i
    
    'FAC_Entête Worksheet
    Dim wsFACEntete As Worksheet: Set wsFACEntete = wshFAC_Entête
    Dim lastUsedRowFacEntete As Long
    lastUsedRowFacEntete = wsFACEntete.Cells(wsFACEntete.Rows.count, 1).End(xlUp).row
    Dim rngFACEntete As Range: Set rngFACEntete = wsFACEntete.Range("A2:A" & lastUsedRowFacEntete)
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshENC_Détails'")
    r = r + 1
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim pmtNo As Long, oldpmtNo As Long
    Dim result As Variant
    'Dictionary pour accumuler les encaissements par facture
    Dim dictENC As Scripting.Dictionary
    Set dictENC = New Scripting.Dictionary
    Dim totalEncDetails As Currency
    
    Dim isEncDétailsValid As Boolean
    isEncDétailsValid = True
    
    For i = 2 To lastUsedRowDetails
        pmtNo = CLng(ws.Cells(i, fEncDPayID).Value)
        If pmtNo <> oldpmtNo Then
            If InStr(strPmtNo, pmtNo) = 0 Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le paiement '" & pmtNo & "' à la ligne " & i & " n'existe pas dans ENC_Entête")
                r = r + 1
                isEncDétailsValid = False
            End If
            strPmtNo = strPmtNo & pmtNo & "|"
            oldpmtNo = pmtNo
        End If
        
        Dim Inv_No As String
        Inv_No = CStr(ws.Cells(i, fEncDInvNo).Value)
        result = Application.WorksheetFunction.XLookup(Inv_No, _
                        rngFACEntete, _
                        rngFACEntete, _
                        "Not Found", _
                        0, _
                        1)
        If result = "Not Found" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La facture '" & Inv_No & "', ligne " & i & ", du paiement '" & pmtNo & "' n'existe pas dans FAC_Entête")
            r = r + 1
            isEncDétailsValid = False
        End If
        
        If IsDate(ws.Cells(i, fEncDPayDate).Value) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La date '" & ws.Cells(i, fEncDPayDate).Value & "', ligne " & i & ", du paiment '" & pmtNo & "' est INVALIDE '")
            r = r + 1
            isEncDétailsValid = False
        End If
        
        If IsNumeric(ws.Cells(i, fEncDPayAmount).Value) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le montant '" & ws.Cells(i, fEncDPayAmount).Value & "' du paiement '" & pmtNo & "' n'est pas numérique")
            r = r + 1
            isEncDétailsValid = False
        Else
            If dictENC.Exists(Inv_No) Then
                dictENC(Inv_No) = dictENC(Inv_No) + ws.Cells(i, fEncDPayAmount).Value
            Else
                dictENC.Add Inv_No, ws.Cells(i, fEncDPayAmount).Value
            End If
            totalEncDetails = totalEncDetails + ws.Cells(i, fEncDPayAmount).Value
        End If
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(lastUsedRowDetails - 1, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 1
    
    'Compare les encaissements accumulés (dictENC) avec wshFAC_Comptes_Clients
    Dim wsComptes_Clients As Worksheet: Set wsComptes_Clients = wshFAC_Comptes_Clients
    Dim lastUsedRow As Long
    lastUsedRow = wsComptes_Clients.Cells(wsComptes_Clients.Rows.count, 1).End(xlUp).row
    Dim totalPaid As Currency
    
    For i = 3 To lastUsedRow
        Inv_No = wsComptes_Clients.Cells(i, fFacCCInvNo).Value
        totalPaid = wsComptes_Clients.Cells(i, fFacCCTotalPaid).Value
        If totalPaid <> dictENC(Inv_No) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Ligne # " & i & " - La facture '" & Inv_No & "', le total des enc. " _
                            & "(wshFAC_Comptes_clients) " & Format$(totalPaid, "###,##0.00 $") _
                            & " est <> du détail des enc. " & Format$(dictENC(Inv_No), "###,##0.00 $"))
            r = r + 1
            isEncDétailsValid = False
        End If
    Next i
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Total des encaissements : " & Format$(totalEncDetails, "##,###,##0.00 $")
    rng.Characters(InStr(rng.Value, Left(totalEncDetails, 1)), 12).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totalEncDetails, 1)), 12).Font.Bold = True

    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + lastUsedRowDetails - 1
    
    'Cas problème dans cette vérification ?
    If isEncDétailsValid = False Then
        verificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set dictENC = Nothing
    Set rng = Nothing
    Set rngEntete = Nothing
    Set rngFACEntete = Nothing
    Set ws = Nothing
    Set wsComptes_Clients = Nothing
    Set wsFACEntete = Nothing
    Set wsEntete = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkENC_Détails", startTime)

End Sub

Private Sub checkENC_Entête(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkENC_Entête", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'Clients Master File
    Dim wsClients As Worksheet: Set wsClients = wshBD_Clients
    Dim lastUsedRowClient As Long
    lastUsedRowClient = wsClients.Cells(wsClients.Rows.count, "B").End(xlUp).row
    Dim rngClients As Range: Set rngClients = wsClients.Range("B2:B" & lastUsedRowClient)
    
    'wshENC_Entête
    Dim ws As Worksheet: Set ws = wshENC_Entête
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lastUsedRow <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshENC_Entête'")
    r = r + 1
    
    If lastUsedRow = HeaderRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim arr As Variant
    arr = wshENC_Entête.Range("A1").CurrentRegion.offset(1, 0) _
              .Resize(lastUsedRow - HeaderRow, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim pmtNo As String
    Dim totals As Currency
    Dim result As Variant
    
    Dim isEncEntêteValid As Boolean
    isEncEntêteValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        pmtNo = arr(i, 1)
        If IsDate(arr(i, 2)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La date de paiement '" & arr(i, 2) & "' du paiement '" & arr(i, 1) & "' n'est pas VALIDE")
            r = r + 1
            isEncEntêteValid = False
        End If
        
        Dim codeClient As String
        codeClient = arr(i, 4)
        result = Application.WorksheetFunction.XLookup(codeClient, _
                        rngClients, _
                        rngClients, _
                        "Not Found", _
                        0, _
                        1)
        If result = "Not Found" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le client '" & codeClient & "' du paiement '" & pmtNo & "' est INVALIDE")
            r = r + 1
            isEncEntêteValid = False
        End If
        totals = totals + arr(i, 6)
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
    r = r + 1
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Total des encaissements : " & Format$(totals, "##,###,##0.00 $")
    rng.Characters(InStr(rng.Value, Left(totals, 1)), 12).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totals, 1)), 12).Font.Bold = True
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    'Cas problème dans cette vérification ?
    If isEncEntêteValid = False Then
        verificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set rng = Nothing
    Set rngClients = Nothing
    Set ws = Nothing
    Set wsClients = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkENC_Entête", startTime)

End Sub

Private Sub checkCC_Régularisations(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkCC_Régularisations", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshCC_Régularisations
    Dim ws As Worksheet: Set ws = wshCC_Régularisations
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRowDetails As Long
    lastUsedRowDetails = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRowDetails <= 2 - HeaderRow Or ws.Cells(lastUsedRowDetails, 1).Value = "" Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRowDetails, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    'FAC_Entête Worksheet
    Dim wsFACEntete As Worksheet: Set wsFACEntete = wshFAC_Entête
    Dim lastUsedRowFacEntete As Long
    lastUsedRowFacEntete = wsFACEntete.Cells(wsFACEntete.Rows.count, 1).End(xlUp).row
    Dim rngFACEntete As Range: Set rngFACEntete = wsFACEntete.Range("A2:A" & lastUsedRowFacEntete)
    
    'Clients Master File
    Dim wsClients As Worksheet: Set wsClients = wshBD_Clients
    Dim lastUsedRowClient As Long
    lastUsedRowClient = wsClients.Cells(wsClients.Rows.count, "B").End(xlUp).row
    Dim rngClients As Range: Set rngClients = wsClients.Range("B2:B" & lastUsedRowClient)
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshCC_Régularisations'")
    r = r + 1
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim regulNo As Long
    Dim result As Variant
    'Dictionary pour accumuler les encaissements par facture
    Dim dictRegul As Scripting.Dictionary
    Set dictRegul = New Scripting.Dictionary
    Dim totalRégularisations As Currency
    
    Dim isRegularisationValid As Boolean
    isRegularisationValid = True
    
    Dim Inv_No As String
    Dim i As Long
    For i = 2 To lastUsedRowDetails
        regulNo = CLng(ws.Cells(i, fREGULRegulID).Value)
        Inv_No = CStr(ws.Cells(i, fREGULInvNo).Value)
        result = Application.WorksheetFunction.XLookup(Inv_No, _
                        rngFACEntete, _
                        rngFACEntete, _
                        "Not Found", _
                        0, _
                        1)
        If result = "Not Found" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La facture '" & Inv_No & "', ligne " & i & ", de la régularisation '" & regulNo & "' n'existe pas dans FAC_Entête")
            r = r + 1
            isRegularisationValid = False
        End If
        
        If IsDate(ws.Cells(i, fREGULDate).Value) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La date '" & ws.Cells(i, fREGULDate).Value & "', ligne " & i & ", de la régularisation '" & regulNo & "' est INVALIDE '")
            r = r + 1
            isRegularisationValid = False
        End If
        
        Dim codeClient As String
        codeClient = ws.Cells(i, fREGULClientID).Value
        result = Application.WorksheetFunction.XLookup(codeClient, _
                        rngClients, _
                        rngClients, _
                        "Not Found", _
                        0, _
                        1)
        If result = "Not Found" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le client '" & codeClient & "' de la régularisation '" & regulNo & "' est INVALIDE")
            r = r + 1
            isRegularisationValid = False
        End If
        
        'Vérification du montant des honoraires
        If IsNumeric(ws.Cells(i, fREGULHono).Value) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le montant des honoraires '" & ws.Cells(i, fREGULHono).Value & "' de la régularisation '" & regulNo & "' n'est pas numérique")
            r = r + 1
            isRegularisationValid = False
        Else
            If dictRegul.Exists(Inv_No) Then
                dictRegul(Inv_No) = dictRegul(Inv_No) + ws.Cells(i, fREGULHono).Value
            Else
                dictRegul.Add Inv_No, ws.Cells(i, fREGULHono).Value
            End If
            totalRégularisations = totalRégularisations + ws.Cells(i, fREGULHono).Value
        End If
        
        'Vérification du montant des frais
        If IsNumeric(ws.Cells(i, fREGULFrais).Value) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le montant des frais '" & ws.Cells(i, fREGULFrais).Value & "' de la régularisation '" & regulNo & "' n'est pas numérique")
            r = r + 1
            isRegularisationValid = False
        Else
            If dictRegul.Exists(Inv_No) Then
                dictRegul(Inv_No) = dictRegul(Inv_No) + ws.Cells(i, fREGULFrais).Value
            Else
                dictRegul.Add Inv_No, ws.Cells(i, fREGULFrais).Value
            End If
            totalRégularisations = totalRégularisations + ws.Cells(i, fREGULFrais).Value
        End If
    
        'Vérification du montant de TPS
        If IsNumeric(ws.Cells(i, fREGULTPS).Value) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le montant de TPS '" & ws.Cells(i, fREGULTPS).Value & "' de la régularisation '" & regulNo & "' n'est pas numérique")
            r = r + 1
            isRegularisationValid = False
        Else
            If dictRegul.Exists(Inv_No) Then
                dictRegul(Inv_No) = dictRegul(Inv_No) + ws.Cells(i, fREGULTPS).Value
            Else
                dictRegul.Add Inv_No, ws.Cells(i, fREGULTPS).Value
            End If
            totalRégularisations = totalRégularisations + ws.Cells(i, fREGULTPS).Value
        End If
        
        'Vérification du montant de TVQ
        If IsNumeric(ws.Cells(i, fREGULTVQ).Value) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le montant de TPS '" & ws.Cells(i, fREGULTVQ).Value & "' de la régularisation '" & regulNo & "' n'est pas numérique")
            r = r + 1
            isRegularisationValid = False
        Else
            If dictRegul.Exists(Inv_No) Then
                dictRegul(Inv_No) = dictRegul(Inv_No) + ws.Cells(i, fREGULTVQ).Value
            Else
                dictRegul.Add Inv_No, ws.Cells(i, fREGULTVQ).Value
            End If
            totalRégularisations = totalRégularisations + ws.Cells(i, fREGULTVQ).Value
        End If
    
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(lastUsedRowDetails - 1, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 1
    
    'Compare les régularisations accumulés (dictRegul) avec wshFAC_Comptes_Clients
    Dim wsComptes_Clients As Worksheet: Set wsComptes_Clients = wshFAC_Comptes_Clients
    Dim lastUsedRow As Long
    lastUsedRow = wsComptes_Clients.Cells(wsComptes_Clients.Rows.count, 1).End(xlUp).row
    Dim totalRegul As Currency
    
    For i = 3 To lastUsedRow
        Inv_No = wsComptes_Clients.Cells(i, fFacCCInvNo).Value
        totalRegul = wsComptes_Clients.Cells(i, fFacCCTotalRegul).Value
        If totalRegul <> dictRegul(Inv_No) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Pour la facture '" & Inv_No & "', le total des régularisations de " _
                            & "(wshFAC_Comptes_clients) " & Format$(totalRegul, "###,##0.00 $") _
                            & " est <> du détail des régularisations (wshCC_Régularisations) " & Format$(dictRegul(Inv_No), "###,##0.00 $"))
            r = r + 1
            isRegularisationValid = False
        End If
    Next i
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Total des régularisations : " & Format$(totalRégularisations, "##,###,##0.00 $")
    rng.Characters(InStr(rng.Value, Left(totalRégularisations, 1)), 12).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totalRégularisations, 1)), 12).Font.Bold = True
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + lastUsedRowDetails - 1
    
    'Cas problème dans cette vérification ?
    If isRegularisationValid = False Then
        verificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set dictRegul = Nothing
    Set rngClients = Nothing
    Set rngFACEntete = Nothing
    Set ws = Nothing
    Set wsClients = Nothing
    Set wsFACEntete = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkCC_Régularisations", startTime)

End Sub

Private Sub checkFAC_Détails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkFAC_Détails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshFAC_Détails
    Dim ws As Worksheet: Set ws = wshFAC_Détails
    Dim HeaderRow As Long: HeaderRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Entête
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A" & 1 + HeaderRow & ":A" & lastUsedRowEntete)
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Détails'")
    r = r + 1
    
    'Transfer FAC_Details data from Worksheet into an Array (arr)
    Dim arr As Variant
    arr = wshFAC_Détails.Range("A1").CurrentRegion.offset(1, 0).Value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String, oldInv_No As String
    Dim result As Variant
    
    Dim isFACDétailsValid As Boolean
    isFACDétailsValid = True
    
    For i = LBound(arr, 1) + 2 To UBound(arr, 1) - 1 'Two lines of header !
        Inv_No = CStr(arr(i, 1))
'        Debug.Print "#018 - Inv_no = ", Inv_No, ", de type ", TypeName(Inv_No)
        If Inv_No <> oldInv_No Then
             result = Application.WorksheetFunction.XLookup(Inv_No, _
                                                    rngMaster, _
                                                    rngMaster, _
                                                    "Not Found", _
                                                    0, _
                                                    1)
            If result = "Not Found" Then
                Debug.Print "#019 " & result
            End If
'            result = Application.WorksheetFunction.XLookup(ws.Cells(i, 1), rngMaster, rngMaster, "Not Found", 0, 1)
            oldInv_No = CStr(Inv_No)
        End If
        If result = "Not Found" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " n'existe pas dans FAC_Entête")
            r = r + 1
            isFACDétailsValid = False
        End If
        If IsNumeric(arr(i, 3)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " le nombre d'heures est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
            isFACDétailsValid = False
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " le taux horaire est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
            isFACDétailsValid = False
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " le montant est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
            isFACDétailsValid = False
        End If
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 2, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - 2
    
    'Cas problème dans cette vérification ?
    If isFACDétailsValid = False Then
        verificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set rngMaster = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkFAC_Détails", startTime)

End Sub

Private Sub checkFAC_Entête(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkFAC_Entête", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshFAC_Entête
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    Dim HeaderRow As Long: HeaderRow = 2
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Entête'")
    r = r + 1
    
    If lastUsedRow = HeaderRow Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    'Prépare à charger les données en mémoire (arr)
    Dim rngData As Range
    Set rngData = ws.Range("A1").CurrentRegion
    Set rngData = rngData.offset(2, 0).Resize(rngData.Rows.count - 2, rngData.Columns.count)
    Dim arr As Variant
    arr = rngData
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String
    Dim totals(1 To 8, 1 To 2) As Currency
    Dim nbFactC As Long, nbFactAC As Long
    
    Dim isFACEntêteValid As Boolean
    isFACEntêteValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        Inv_No = arr(i, 1)
        If IsDate(arr(i, 2)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
            isFACEntêteValid = False
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La facture '" & Inv_No & "' à la ligne " & i & ", la date est de mauvais format '" & arr(i, 2) & "'")
                r = r + 1
                isFACEntêteValid = False
            End If
        End If
        If arr(i, 3) <> "C" And arr(i, 3) <> "AC" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le type de facture '" & arr(i, 3) & "' pour la facture '" & Inv_No & "' est INVALIDE")
            isFACEntêteValid = False
            r = r + 1
        End If
        If arr(i, 19) <> 0.09975 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le % de TVQ, soit '" & arr(i, 19) & "' pour la facture '" & Inv_No & "' est ERRONÉ")
            r = r + 1
            isFACEntêteValid = False
        End If
        If arr(i, 3) = "C" Then
            totals(1, 1) = totals(1, 1) + arr(i, 10)
            totals(2, 1) = totals(2, 1) + arr(i, 12)
            totals(3, 1) = totals(3, 1) + arr(i, 14)
            totals(4, 1) = totals(4, 1) + arr(i, 16)
            totals(5, 1) = totals(5, 1) + arr(i, 18)
            totals(6, 1) = totals(6, 1) + arr(i, 20)
            totals(7, 1) = totals(7, 1) + arr(i, 21)
            totals(8, 1) = totals(8, 1) + arr(i, 22)
            nbFactC = nbFactC + 1
        Else
            totals(1, 2) = totals(1, 2) + arr(i, 10)
            totals(2, 2) = totals(2, 2) + arr(i, 12)
            totals(3, 2) = totals(3, 2) + arr(i, 14)
            totals(4, 2) = totals(4, 2) + arr(i, 16)
            totals(5, 2) = totals(5, 2) + arr(i, 18)
            totals(6, 2) = totals(6, 2) + arr(i, 20)
            totals(7, 2) = totals(7, 2) + arr(i, 21)
            totals(8, 2) = totals(8, 2) + arr(i, 22)
            nbFactAC = nbFactAC + 1
        End If
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
    r = r + 1
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Totaux des factures CONFIRMÉES (" & nbFactC & " factures)"
    rng.Characters(InStr(rng.Value, "CONFIRMÉES"), 10).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, "CONFIRMÉES"), 10).Font.Bold = True
    r = r + 1

    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Total Fact. : " & Fn_Pad_A_String(Format$(totals(7, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left(totals(7, 1), 1)), 15).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totals(7, 1), 1)), 15).Font.Bold = True
    r = r + 1

    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Acompte payé: " & _
            Fn_Pad_A_String(Format$(totals(8, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 2
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "Totaux des factures À CONFIRMER (" & nbFactAC & " factures)"
    rng.Characters(InStr(rng.Value, "À CONFIRMER"), 11).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, "À CONFIRMER"), 11).Font.Bold = True
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Total Fact. : " & Fn_Pad_A_String(Format$(totals(7, 2), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left(totals(7, 2), 1)), 15).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totals(7, 2), 1)), 15).Font.Bold = True
    r = r + 1

    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Acompte payé: " & _
            Fn_Pad_A_String(Format$(totals(8, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - HeaderRow
    
    'Cas problème dans cette vérification ?
    If isFACEntêteValid = False Then
        verificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set rng = Nothing
    Set rngData = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkFAC_Entête", startTime)

End Sub

Private Sub checkFAC_Comptes_Clients(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkFAC_Comptes_Clients", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
    Dim HeaderRow As Long: HeaderRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Comptes_Clients'")
    r = r + 1
    
    If lastUsedRow = HeaderRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    'Load every records into an Array
    Dim arr As Variant
    arr = wshFAC_Comptes_Clients.Range("A1").CurrentRegion.offset(2, 0) _
              .Resize(lastUsedRow - HeaderRow, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String
    Dim totals(1 To 4, 1 To 2) As Currency
    Dim nbFactC As Long, nbFactAC As Long
    
    Dim isFACCCValid As Boolean
    isFACCCValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        Inv_No = arr(i, fFacCCInvNo)
        Dim invType As String
        invType = Fn_Get_Invoice_Type(Inv_No)
        If invType <> "C" And invType <> "AC" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le type de facture '" & invType & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Date ?
        If IsDate(CDate(arr(i, fFacCCInvoiceDate))) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", la date '" & arr(i, fFacCCInvoiceDate) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        Else
            If arr(i, fFacCCInvoiceDate) <> Int(arr(i, fFacCCInvoiceDate)) Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", la facture '" & Inv_No & "', la date est de mauvais format '" & arr(i, fFacCCInvoiceDate) & "'")
                r = r + 1
                isFACCCValid = False
            End If
        End If
        
        'Code client ?
        If Fn_Validate_Client_Number(CStr(arr(i, fFacCCCodeClient))) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le client '" & CStr(arr(i, fFacCCCodeClient)) & "' de la facture '" & Inv_No & "' est INVALIDE '")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Status (Paid or Unpaid)
        If arr(i, fFacCCStatus) <> "Paid" And arr(i, fFacCCStatus) <> "Unpaid" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le statut '" & arr(i, fFacCCStatus) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Date due
        If IsDate(CDate(arr(i, fFacCCDueDate))) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", la date due '" & arr(i, fFacCCDueDate) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Total
        If IsNumeric(arr(i, fFacCCTotal)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le total de la facture '" & arr(i, fFacCCTotal) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Encaissé à date
        If IsNumeric(arr(i, fFacCCTotalPaid)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le montant payé à date '" & arr(i, fFacCCTotalPaid) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Régularisation à date
        If IsNumeric(arr(i, fFacCCTotalRegul)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le montant régularisé à date '" & arr(i, fFacCCTotalRegul) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Solde à recevoir
        If IsNumeric(arr(i, fFacCCBalance)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", le solde de la facture '" & arr(i, fFacCCBalance) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        
        'PLUG pour s'assurer que le solde impayé est bel et bien aligné sur le total moins le payé à date - les régularisations
'        If arr(i, fFacCCBalance) <> arr(i, fFacCCTotal) - arr(i, fFacCCTotalPaid) Then
        If arr(i, fFacCCBalance) <> arr(i, fFacCCTotal) - arr(i, fFacCCTotalPaid) + arr(i, fFacCCTotalRegul) Then
'            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", pour la facture '" & Inv_No & _
'                    ", le solde à recevoir est erroné, soit " & Format$(arr(i, fFacCCBalance), "###,##0.00 $") & "' vs. " & _
'                    "'" & Format$(arr(i, fFacCCTotal) - arr(i, fFacCCTotalPaid), "###,##0.00 $") & "'")
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + 2 & ", pour la facture '" & Inv_No & _
                    ", le solde à recevoir est erroné, soit " & Format$(CCur(arr(i, fFacCCBalance)), "###,##0.00 $") & "' vs. " & _
                    "'" & Format$(arr(i, fFacCCTotal) - arr(i, fFacCCTotalPaid) + arr(i, fFacCCTotalRegul), "###,##0.00 $") & "'")
            r = r + 1
            isFACCCValid = False
        End If
        
        'Le statut est-il exact
        If arr(i, fFacCCBalance) = 0 And arr(i, fFacCCStatus) = "Unpaid" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le statut '" & arr(i, fFacCCStatus) & "' de la facture '" & Inv_No & "', avec un solde de " & Format$(arr(i, fFacCCBalance), "#,##0.00 $") & " est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        If arr(i, fFacCCBalance) <> 0 And arr(i, fFacCCStatus) = "Paid" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le statut '" & arr(i, fFacCCStatus) & "' de la facture '" & Inv_No & "', avec un solde de " & Format$(arr(i, fFacCCBalance), "#,##0.00 $") & " est INVALIDE")
            r = r + 1
            isFACCCValid = False
        End If
        If invType = "C" Then
            totals(1, 1) = totals(1, 1) + arr(i, fFacCCTotal)
            totals(2, 1) = totals(2, 1) + arr(i, fFacCCTotalPaid)
            totals(3, 1) = totals(3, 1) + arr(i, fFacCCTotalRegul)
            totals(4, 1) = totals(4, 1) + arr(i, fFacCCBalance)
            nbFactC = nbFactC + 1
        Else
            totals(1, 2) = totals(1, 2) + arr(i, fFacCCTotal)
            totals(2, 2) = totals(2, 2) + arr(i, fFacCCTotalPaid)
            totals(3, 2) = totals(3, 2) + arr(i, fFacCCTotalRegul)
           totals(4, 2) = totals(4, 2) + arr(i, fFacCCBalance)
            nbFactAC = nbFactAC + 1
        End If
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
    r = r + 1
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "Totaux des factures CONFIRMÉES (" & nbFactC & " factures)"
    rng.Characters(InStr(rng.Value, "CONFIRMÉES"), 10).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, "CONFIRMÉES"), 10).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Total des factures         : " & Fn_Pad_A_String(Format$(totals(1, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left(totals(1, 1), 1)), 15).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totals(1, 1), 1)), 15).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Montants encaissés à date  : " & Fn_Pad_A_String(Format$(totals(2, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left(totals(2, 1), 1)), 15).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totals(2, 1), 1)), 15).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Montant régularisé à date  : " & Fn_Pad_A_String(Format$(totals(3, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left(totals(3, 1), 1)), 15).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totals(3, 1), 1)), 15).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Solde à recevoir           : " & Fn_Pad_A_String(Format$(totals(4, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left(totals(4, 1), 1)), 15).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totals(4, 1), 1)), 15).Font.Bold = True
    r = r + 2
    soldeComptesClients = totals(4, 1)
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "Total des factures À CONFIRMER (" & nbFactAC & " factures)"
    rng.Characters(InStr(rng.Value, "À CONFIRMER"), 11).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, "À CONFIRMER"), 11).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Total des factures        : " & Fn_Pad_A_String(Format$(totals(1, 2), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.Value, Left(totals(1, 2), 1)), 15).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(totals(1, 2), 1)), 15).Font.Bold = True
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - HeaderRow
    
    'Cas problème dans cette vérification ?
    If isFACCCValid = False Then
        verificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkFAC_Comptes_Clients", startTime)

End Sub

Private Sub checkFAC_Sommaire_Taux(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkFAC_Sommaire_Taux", 0)

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Sommaire_Taux
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    Dim tol As Double
    tol = 0.0001 'Petite tolérance pour la comparaison
    
    'wshFAC_Sommaire_Taux
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= 2 - HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Sommaire_Taux'")
    r = r + 1
    
    'On a besoin des factures pour la validation
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Entête
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A2:A" & lastUsedRowEntete)
    
    'On a besoin des professionnels
    Dim rngProf As Range
    Call Get_Range_From_Dynamic_Named_Range("dnrProf_Initials_Only", rngProf)

    'Copie les données vers un tableau
    Dim rng As Range
    Set rng = ws.Range("A1:E" & lastUsedRow)
    Dim arr() As Variant
    
    Dim isFACSTValid As Boolean
    isFACSTValid = True
    
    Dim headerRows As Long
    headerRows = 1
    Call Tx_Range_2_2D_Array(rng, arr, 1)
    
    'On analyse chacune des lignes du tableau
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Fn_Is_String_Valid(CStr(arr(i, 1)), rngMaster) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la facture '" & arr(i, 1) & "' est INVALIDE")
            r = r + 1
            isFACSTValid = False
        End If
        If IsNumeric(arr(i, 2)) = False Or arr(i, 2) <> CLng(Int(arr(i, 2))) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la valeur de la séquence '" & arr(i, 2) & "' est INVALIDE")
            r = r + 1
            isFACSTValid = False
        End If
        
        If Fn_Is_String_Valid(CStr(arr(i, 3)), rngProf) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le professionnel '" & arr(i, 3) & "' est INVALIDE")
            r = r + 1
            isFACSTValid = False
        End If
        
        'Heures
        If IsNumeric(arr(i, 4)) = False Or Abs((arr(i, 4) * 100) - Round(arr(i, 4) * 100, 0)) > tol Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", les heures '" & arr(i, 4) & "' sont INVALIDES")
            r = r + 1
            isFACSTValid = False
        End If
        
        'Taux Horaire
        If IsNumeric(arr(i, 5)) = False Or Abs((arr(i, 5) * 100) - Round(arr(i, 5) * 100, 0)) > tol Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le taux horaire '" & arr(i, 5) & "' est INVALIDE")
            r = r + 1
            isFACSTValid = False
        End If
        readRows = readRows + 1
    Next i

    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Cas problème dans cette vérification ?
    If isFACSTValid = False Then
        verificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set rng = Nothing
    Set rngMaster = Nothing
    Set rngProf = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    Set wsOutput = Nothing

    Application.ScreenUpdating = True

    Call Log_Record("modAppli:checkFAC_Sommaire_Taux", startTime)

End Sub

Private Sub checkFAC_Projets_Entête(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkFAC_Projets_Entête", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Projets_Entête'")
    r = r + 1
    
    'Establish the number of rows before transferring it to an Array
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.Rows.count
    If numRows <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.offset(1, 0).Resize(numRows - 1, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim projetID As String
    Dim codeClient As String
    
    Dim isFacProjetEntêteValid As Boolean
    isFacProjetEntêteValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1) 'One line of header !
        projetID = arr(i, 1)
        'Client valide ?
        codeClient = Trim(arr(i, 3))
        If Fn_Validate_Client_Number(codeClient) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le Code de Client est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If IsDate(arr(i, 4)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 4) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le total des honoraires est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If IsNumeric(arr(i, 7)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du premier sommaire sont INVALIDES '" & arr(i, 7) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du premier sommaire est INVALIDE '" & arr(i, 8) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If IsNumeric(arr(i, 9)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du premier sommaire sont INVALIDES '" & arr(i, 9) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 11) <> "" And IsNumeric(arr(i, 11)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du second sommaire sont INVALIDES '" & arr(i, 11) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 12) <> "" And IsNumeric(arr(i, 12)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du second sommaire est INVALIDE '" & arr(i, 12) & "'")
            r = r + 1
        End If
        If arr(i, 13) <> "" And IsNumeric(arr(i, 13)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du second sommaire sont INVALIDES '" & arr(i, 13) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 15) <> "" And IsNumeric(arr(i, 15)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du troisième sommaire sont INVALIDES '" & arr(i, 15) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 16) <> "" And IsNumeric(arr(i, 16)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du troisième sommaire est INVALIDE '" & arr(i, 16) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 17) <> "" And IsNumeric(arr(i, 17)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du troisième sommaire sont INVALIDES '" & arr(i, 17) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 19) <> "" And IsNumeric(arr(i, 19)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du quatrième sommaire sont INVALIDES '" & arr(i, 19) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 20) <> "" And IsNumeric(arr(i, 20)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du quatrième sommaire est INVALIDE '" & arr(i, 20) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 21) <> "" And IsNumeric(arr(i, 21)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du quatrième sommaire sont INVALIDES '" & arr(i, 21) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 23) <> "" And IsNumeric(arr(i, 23)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du cinquième sommaire sont INVALIDES '" & arr(i, 23) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 24) <> "" And IsNumeric(arr(i, 24)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du cinquième sommaire est INVALIDE '" & arr(i, 24) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
        If arr(i, 25) <> "" And IsNumeric(arr(i, 25)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du cinquième sommaire sont INVALIDES '" & arr(i, 25) & "'")
            r = r + 1
            isFacProjetEntêteValid = False
        End If
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " projets de factures a été analysés")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    'Cas problème dans cette vérification ?
    If isFacProjetEntêteValid = False Then
        verificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkFAC_Projets_Entête", startTime)

End Sub

Private Sub checkFAC_Projets_Détails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkFAC_Projets_Détails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshFAC_Projets_Détails
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If

    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Projets_Entête
    lastUsedRow = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A2:A" & lastUsedRow)
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Projets_Détails'")
    r = r + 1
    
    'Transfer data from Worksheet into an Array (arr)
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.Rows.count - 1 'Remove header
    If numRows < 1 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    'Charge le contenu de 'wshFAC_Projets_Détails' en mémoire (Array)
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.offset(1, 0).Resize(numRows, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim projetID As Long, oldProjetID As Long
    Dim codeClient As String
    Dim result As Variant
    
    'Sauvegarde la ligne active
    Dim saveR As Long
    saveR = r
    
    'À partir de la mémoire (Array)
    For i = LBound(arr, 1) To UBound(arr, 1)
        projetID = CLng(arr(i, 1))
        If projetID <> oldProjetID Then
            result = Application.WorksheetFunction.XLookup(projetID, _
                                rngMaster, _
                                rngMaster, _
                                "Not Found", _
                                0, _
                                1)
            oldProjetID = projetID
        End If
        If result = "Not Found" Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le projet '" & projetID & "' à la ligne " & i & " n'existe pas dans FAC_Projets_Entête")
            r = r + 1
        End If
        'Client valide ?
        codeClient = Trim(arr(i, 3))
        If Fn_Validate_Client_Number(codeClient) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Dans le projet '" & projetID & "' à la ligne " & i & " le Code de Client est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le projet '" & projetID & "' à la ligne " & i & " le TECID est INVALIDE '" & arr(i, 4) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le projet '" & projetID & "' à la ligne " & i & " le ProfID est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le projet '" & projetID & "' à la ligne " & i & " les Heures sont INVALIDES '" & arr(i, 8) & "'")
            r = r + 1
        End If
    Next i
    
    'Est-ce qu'il y a eu des messages de générés ?
    If saveR <> r Then
        verificationIntegriteOK = False
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes ont été analysées")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - HeaderRow
    
Clean_Exit:
    'Libérer la mémoire
    Set rngMaster = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkFAC_Projets_Détails", startTime)

End Sub

Private Sub checkGL_Trans(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkGL_Trans", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshGL_Trans
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lastUsedRow <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim firstEmptyCol As Long
    firstEmptyCol = 1
    Do Until ws.Cells(HeaderRow, firstEmptyCol) = ""
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshGL_Trans'")
    r = r + 1
    
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wshAdmin.Range("dnrPlanComptable_All")
    On Error GoTo 0

    Dim isGLTransValid As Boolean
    isGLTransValid = True
    
    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        r = r + 1
        isGLTransValid = False
        Exit Sub
    End If
    
    Dim strCodeGL As String, strDescGL As String
    Dim ligne As Range
    For Each ligne In planComptable.Rows
        strCodeGL = strCodeGL & ligne.Cells(1, 2).Value & "|:|"
        strDescGL = strDescGL & ligne.Cells(1, 1).Value & "|:|"
    Next ligne
    
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.Rows.count - 1 'Remove the header row
    If numRows < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.offset(1, 0).Resize(numRows, ws.Range("A1").CurrentRegion.Columns.count).Value
    
    Dim dict_GL_Entry As New Dictionary
    Dim sum_arr() As Double
    ReDim sum_arr(1 To 2500, 1 To 3)
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim dt As Currency, ct As Currency
    Dim arTotal As Currency
    Dim GL_Entry_No As String, glCode As String, glDescr As String
    Dim result As Variant
    Dim CCGlNo As String
    CCGlNo = ObtenirNoGlIndicateur("Comptes Clients")
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        GL_Entry_No = arr(i, 1)
        If dict_GL_Entry.Exists(GL_Entry_No) = False Then
            dict_GL_Entry.Add GL_Entry_No, row
            sum_arr(row, 1) = GL_Entry_No
            row = row + 1
        End If
        If IsDate(arr(i, 2)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** L'écriture #  " & GL_Entry_No & " ' à la ligne " & i & " a une date INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
            isGLTransValid = False
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** L'écriture #  " & GL_Entry_No & " ' à la ligne " & i & " a une date avec le mauvais format '" & arr(i, 2) & "'")
                r = r + 1
                isGLTransValid = False
            End If
        End If
        glCode = CStr(arr(i, 5))
        If InStr(1, strCodeGL, glCode + "|:|") = 0 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le compte '" & glCode & "' à la ligne " & i & " est INVALIDE '")
            r = r + 1
            isGLTransValid = False
        End If
        If glCode = CCGlNo Then
            arTotal = arTotal + arr(i, 7) - arr(i, 8)
        End If
        glDescr = arr(i, 6)
        If InStr(1, strDescGL, glDescr + "|:|") = 0 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La description du compte '" & glDescr & "' à la ligne " & i & " est INVALIDE")
            r = r + 1
            isGLTransValid = False
        End If
        dt = arr(i, 7)
        If IsNumeric(dt) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le montant du débit '" & dt & "' à la ligne " & i & " n'est pas une valeur numérique")
            r = r + 1
            isGLTransValid = False
        End If
        ct = arr(i, 8)
        If IsNumeric(ct) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le montant du débit '" & ct & "' à la ligne " & i & " n'est pas une valeur numérique")
            r = r + 1
            isGLTransValid = False
        End If
        currentRow = dict_GL_Entry(GL_Entry_No)
        sum_arr(currentRow, 2) = sum_arr(currentRow, 2) + dt
        sum_arr(currentRow, 3) = sum_arr(currentRow, 3) + ct
        If arr(i, 10) <> "" Then
            If IsDate(arr(i, 10)) = False Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le TimeStamp '" & arr(i, 10) & "' à la ligne " & i & " n'est pas une date VALIDE")
                r = r + 1
                isGLTransValid = False
            End If
        End If
    Next i
    
    Dim sum_dt As Currency, sum_ct As Currency
    Dim cas_hors_balance As Long
    Dim v As Variant
    For Each v In dict_GL_Entry.items()
        GL_Entry_No = sum_arr(v, 1)
        dt = Round(sum_arr(v, 2), 2)
        ct = Round(sum_arr(v, 3), 2)
        If dt <> ct Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Écriture # " & v & " ne balance pas... Dt = " & Format$(dt, "###,###,##0.00") & " et Ct = " & Format$(ct, "###,###,##0.00"))
            r = r + 1
            isGLTransValid = False
            cas_hors_balance = cas_hors_balance + 1
        End If
        sum_dt = sum_dt + dt
        sum_ct = sum_ct + ct
    Next v
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - HeaderRow, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - HeaderRow
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Un total de " & dict_GL_Entry.count & " écritures ont été analysées")
    r = r + 1
    
    If cas_hors_balance = 0 Then
        'Un peu de couleur
        Dim rng As Range: Set rng = wsOutput.Range("B" & r)
        rng.Value = "       Chacune des écritures balancent au niveau de l'écriture"
        rng.Characters(InStr(rng.Value, "C"), Len(rng.Value) - 7).Font.COLOR = vbRed
        rng.Characters(InStr(rng.Value, "C"), Len(rng.Value) - 7).Font.Bold = True
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_hors_balance & " écriture(s) qui ne balance(nt) pas !!!")
        r = r + 1
        isGLTransValid = False
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Les totaux des transactions sont:")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Dt = " & Format$(sum_dt, "###,###,##0.00 $"))
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Ct = " & Format$(sum_ct, "###,###,##0.00 $"))
    r = r + 1
    
    If sum_dt - sum_ct <> 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Hors-Balance de " & Format$(sum_dt - sum_ct, "###,###,##0.00 $"))
        r = r + 1
        isGLTransValid = False
    End If
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.Value = "Au Grand Livre, le solde des Comptes-Clients est de : " & Format$(arTotal, "##,###,##0.00 $")
    rng.Characters(InStr(rng.Value, Left(arTotal, 1)), 15).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, Left(arTotal, 1)), 15).Font.Bold = True
    r = r + 2
    If soldeComptesClients <> arTotal Then
        MsgBox "ATTENTION, le solde des Comptes-Clients" & vbNewLine & vbNewLine & _
                "diffère entre les 2 sources...", vbCritical, "FAC_Comptes_Clients <> Solde au Grand-Livre !!!"
    End If
    
    'Cas problème dans cette vérification ?
    If isGLTransValid = False Then
        verificationIntegriteOK = False
    End If
    
Clean_Exit:
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set ligne = Nothing
    Set planComptable = Nothing
    Set rng = Nothing
    Set v = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Call Log_Record("modAppli:checkGL_Trans", startTime)

End Sub

Private Sub checkGL_EJ_Recurrente(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkGL_EJ_Recurrente", 0)

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshGL_EJ_Recurrente
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshGL_EJ_Recurrente
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= 2 - HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshGL_EJ_Recurrente'")
    r = r + 1
    
    Dim isGlEjRécurrenteValid As Boolean
    isGlEjRécurrenteValid = True
    
    'On a besoin des comptes du G/L pour la validation
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wshAdmin.Range("dnrPlanComptable_All")
    On Error GoTo 0

    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation, "modAppli_Utils:checkGL_EJ_Recurrente"
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        r = r + 1
        isGlEjRécurrenteValid = False
        Exit Sub
    End If
    
    'Bâtir une chaine avec code & description
    Dim strGL As String
    Dim ligne As Range
    For Each ligne In planComptable.Rows
        strGL = strGL & Trim(ligne.Cells(1, 2).Value) & "-" & Trim(ligne.Cells(1, 1).Value) & " | "
    Next ligne

    'Copier les données vers un tableau
    Dim rng As Range
    Set rng = ws.Range("A1:G" & lastUsedRow)
    Dim arr() As Variant
    Dim headerRows As Long
    headerRows = 1
    Call Tx_Range_2_2D_Array(rng, arr, 1)
    
    'On analyse chacune des lignes du tableau
    Dim i As Long, p As Long
    Dim GL As String, descGL As String
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsNumeric(arr(i, 1)) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le numéro d'écriture '" & arr(i, 1) & "' est INVALIDE")
            r = r + 1
            isGlEjRécurrenteValid = False
        End If
        
        p = InStr(strGL, Trim(arr(i, 3)))
        If p = 0 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le poste de G/L '" & arr(i, 3) & "' est INVALIDE")
            r = r + 1
            isGlEjRécurrenteValid = False
        End If
        If p > 0 Then
            GL = Mid(strGL, p)
            descGL = Mid(GL, InStr(GL, "-") + 1, InStr(GL, " | ") - 6)
            If descGL <> Trim(arr(i, 4)) Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", la description du G/L '" & arr(i, 4) & "' est INVALIDE")
                r = r + 1
                isGlEjRécurrenteValid = False
            End If
        End If
        If arr(i, 5) <> "" Then
            If IsNumeric(arr(i, 5)) = False Or arr(i, 5) <> Round(arr(i, 5), 2) Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le $ du débit '" & arr(i, 5) & "' est INVALIDE")
                r = r + 1
                isGlEjRécurrenteValid = False
            End If
        End If
        If arr(i, 6) <> "" Then
            If IsNumeric(arr(i, 6)) = False Or arr(i, 6) <> Round(arr(i, 6), 2) Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le $ du crédit '" & arr(i, 6) & "' est INVALIDE")
                r = r + 1
                isGlEjRécurrenteValid = False
            End If
        End If
        readRows = readRows + 1
    Next i

    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Cas problème dans cette vérification ?
    If isGlEjRécurrenteValid = False Then
        verificationIntegriteOK = False
    End If
    
Clean_Exit:
    'Libérer la mémoire
    Set ligne = Nothing
    Set planComptable = Nothing
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing

    Application.ScreenUpdating = True

    Call Log_Record("modAppli:checkGL_EJ_Recurrente", startTime)

End Sub

Private Sub checkTEC_TdB_Data(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkTEC_TdB_Data", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshTEC_TdB_Data
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    Dim HeaderRow As Long: HeaderRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshTEC_TdB_Data'")
    r = r + 1
    
    'On a besoin des professionnels
    Dim rngProf As Range
    Call Get_Range_From_Dynamic_Named_Range("dnrProf_Initials_Only", rngProf)

    'Copie les données vers un tableau - 2024-11-20 @ 14:19
    Dim rngData As Range
    Set rngData = ws.Range("A1").CurrentRegion
    Dim arr() As Variant
    Dim headerRows As Long: headerRows = 1
    Call Tx_Range_2_2D_Array(rngData, arr, 1)
    
'    Dim arr As Variant
'    arr = ws.Range("A1").CurrentRegion.Offset(1)
    Dim dict_TECID As New Dictionary
    
    Dim i As Long, tecID As Long, profID As String, prof As String, dateTEC As Date, clientCode As String
    Dim minDate As Date, maxDate As Date
    Dim hres As Double, hres_non_detruites As Double
    Dim estDetruit As Boolean, estFacturable As Boolean, estFacturee As Boolean
    Dim cas_doublon_TECID As Long, cas_date_invalide As Long, cas_doublon_prof As Long, cas_doublon_client As Long
    Dim cas_hres_invalide As Long, cas_estFacturable_invalide As Long, cas_estFacturee_invalide As Long
    Dim cas_estDetruit_invalide As Long
    Dim total_hres_inscrites As Double, total_hres_detruites As Double, total_hres_facturees As Double
    Dim total_hres_facturable As Double, total_hres_TEC As Double, total_hres_non_facturable As Double
    
    minDate = "12/31/2999"
    
    Dim isTECTDBValid As Boolean
    isTECTDBValid = True
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        tecID = arr(i, 1)
'        prof = arr(i, 3)
        If Fn_Is_String_Valid(CStr(arr(i, 3)), rngProf) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** À la ligne " & i + headerRows & ", le professionnel '" & arr(i, 3) & "' est INVALIDE")
            r = r + 1
            isTECTDBValid = False
        End If
        dateTEC = arr(i, 4)
        If IsDate(dateTEC) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID =" & tecID & " a une date INVALIDE '" & dateTEC & " !!!")
            r = r + 1
            isTECTDBValid = False
            cas_date_invalide = cas_date_invalide + 1
        Else
            If dateTEC < minDate Then minDate = dateTEC
            If dateTEC > maxDate Then maxDate = dateTEC
        End If
        If dateTEC <> Int(dateTEC) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La date du TEC '" & dateTEC & "' n'est pas du bon format (H:M:S) pour le TECID =" & tecID)
            r = r + 1
            isTECTDBValid = False
        End If
        clientCode = arr(i, 5)
        hres = arr(i, 8)
        If IsNumeric(hres) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            r = r + 1
            isTECTDBValid = False
            cas_hres_invalide = cas_hres_invalide + 1
        End If
        estFacturable = arr(i, 9)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            r = r + 1
            isTECTDBValid = False
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If
        estFacturee = arr(i, 10)
        If InStr("Vrai^Faux^", estFacturee & "^") = 0 Or Len(estFacturee) <> 2 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            r = r + 1
            isTECTDBValid = False
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        estDetruit = arr(i, 11)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            r = r + 1
            isTECTDBValid = False
            cas_estDetruit_invalide = cas_estDetruit_invalide + 1
        End If
        
        'Heures Inscrites
        total_hres_inscrites = total_hres_inscrites + hres
        hres_non_detruites = hres
        
        'Heures détruites
        If estDetruit = "Vrai" Then
            total_hres_detruites = total_hres_detruites + hres
            hres_non_detruites = hres_non_detruites - hres
        End If
        
        'Heures FACTURABLES
        If hres_non_detruites <> 0 And estFacturable = "Vrai" And _
            Fn_Is_Client_Facturable(clientCode) = True Then
                total_hres_facturable = total_hres_facturable + hres_non_detruites
        End If
        
        'Heures non-FACTURABLES
        If hres_non_detruites <> 0 Then
            If estFacturable = "Faux" Or Fn_Is_Client_Facturable(clientCode) = False Then
                total_hres_non_facturable = total_hres_non_facturable + hres_non_detruites
            End If
        End If
        
        'Heures FACTURÉES
        If hres_non_detruites <> 0 And estDetruit = "Faux" And estFacturee = "Vrai" And _
            Fn_Is_Client_Facturable(clientCode) = True Then
                total_hres_facturees = total_hres_facturees + hres_non_detruites
        End If
        
        'Dictionary
        If dict_TECID.Exists(tecID) = False Then
            dict_TECID.Add tecID, 0
        Else
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le TECID '" & tecID & "' est un doublon pour la ligne '" & i & "'")
            r = r + 1
            isTECTDBValid = False
           cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - HeaderRow, "##,##0") & " charges de temps ont été analysées!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - HeaderRow
    
    If cas_doublon_TECID = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucun doublon de TECID")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_doublon_TECID & " cas de doublons pour les TECID")
        r = r + 1
    End If
    
    If cas_date_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune date INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        r = r + 1
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    r = r + 1
    
    If cas_hres_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune heures INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturable_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune valeur 'estFacturable' n'est INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturee_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune valeur 'estFacturee' n'est INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        r = r + 1
    End If
    
    If cas_estDetruit_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune valeur 'estDetruit' n'est INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        r = r + 1
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "La somme des heures saisies donne ces résultats:")
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Heures SAISIES         : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Heures détruites       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Heures NETTES          : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "              Non_facturables : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "              Facturables     : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Heures facturées       : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable - total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Heures TEC             : " & formattedHours
    rng.Characters(InStr(rng.Value, ":") + 2, Len(formattedHours)).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, ":") + 2, Len(formattedHours)).Font.Bold = True
    r = r + 2

    'Cas problème dans cette vérification ?
    If isTECTDBValid = False Then
        verificationIntegriteOK = False
    End If

Clean_Exit:
    'Libérer la mémoire
    Set rng = Nothing
    Set rngData = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkTEC_TdB_Data", startTime)

End Sub

Private Sub checkTEC(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:checkTEC", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshTEC_Local
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    Dim lastTECIDReported As Long
    lastTECIDReported = 3459 'What is the last TECID analyzed ?

    'Feuille contenant les données à analyser
    Dim HeaderRow As Long: HeaderRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastUsedRow <= HeaderRow Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow, "###,##0") & _
        " lignes et " & Format$(ws.Range("A1").CurrentRegion.Columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshTEC_Local'")
    r = r + 1
    
    'Identifier la plage de données dans wshTEC_Local (exclus les entêtes)
    Dim rngTEC_LocalData As Range
    Set rngTEC_LocalData = ws.Range("A1").CurrentRegion.offset(2, 0)
    Set rngTEC_LocalData = rngTEC_LocalData.Resize(rngTEC_LocalData.Rows.count - 2, rngTEC_LocalData.Columns.count)
    
    'Charger les données dans un tableau (arrTEC_LocalData)
    Dim arrTEC_Local_Data As Variant
    arrTEC_Local_Data = rngTEC_LocalData.Resize(rngTEC_LocalData.Rows.count, rngTEC_LocalData.Columns.count).Value
    
    'Créer un dictionary pour tous les clients
    Dim dictClient As New Dictionary
    'Définir la feuille de calcul et la plage de données
    Dim rngClient As Range
    Set rngClient = wshBD_Clients.Range("A1").CurrentRegion
    Set rngClient = rngClient.offset(1, 0).Resize(rngClient.Rows.count - 1, 2)

    ' Charger les données dans un tableau (deux premières colonnes seulement)
    Dim arr As Variant
    arr = rngClient.Value 'Charger les colonnes 1 et 2

    ' Créer un dictionnaire
    Set dictClient = CreateObject("Scripting.Dictionary")

    ' Remplir le dictionnaire avec les données (clé = colonne 1, valeur = colonne 2)
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        dictClient(arr(i, 2)) = arr(i, 1) ' Utilise la colonne 1 comme clé et la colonne 2 comme valeur
    Next i
    
    'Obtenir toutes les factures émises (wshFAC_Entête) et utiliser un dictionary pour les mémoriser
    Dim rngFAC_EntêteData As Range
    Set rngFAC_EntêteData = wshFAC_Entête.Range("A1").CurrentRegion
    Set rngFAC_EntêteData = rngFAC_EntêteData.offset(2, 0).Resize(rngFAC_EntêteData.Rows.count - 2)
    
    'Calculer le nombre de lignes dans la plage
    Dim lastRow As Long
    lastRow = rngFAC_EntêteData.Rows.count

    ' Redimensionner le tableau pour contenir les données de 2 colonnes
'    Dim arr() As Variant
    ReDim arr(1 To lastRow, 1 To 2)

    ' Remplir le tableau avec les valeurs des colonnes 1 et 3
    For i = 1 To lastRow
        arr(i, 1) = rngFAC_EntêteData.Cells(i, 1).Value
        arr(i, 2) = rngFAC_EntêteData.Cells(i, 3).Value
    Next i

    'Charger dans le dictionnaire dictFacture
    Dim dictFacture As New Dictionary
    On Error Resume Next 'Empêcher les erreurs si doublons
    For i = 1 To UBound(arr, 1)
        dictFacture.Add key:=arr(i, 1), item:=arr(i, 2)
    Next i
    On Error GoTo 0 'Réactiver la gestion normale des erreurs
    
    Dim dictFactureHres As New Dictionary
    Dim dict_TECID As New Dictionary
    Dim dict_prof As New Dictionary
    
    Dim tecID As Long, profID As String, prof As String, dateTEC As Date, dateFact As Date, testDate As Boolean
    Dim minDate As Date, maxDate As Date
    Dim maxTECID As Long
    Dim d As Integer, m As Integer, Y As Integer, p As Integer
    Dim codeClient As String, nomCLient As String
    Dim isClientValid As Boolean
    Dim hres As Double, testHres As Boolean, estFacturable As Boolean
    Dim estFacturee As Boolean, estDetruit As Boolean
    Dim invNo As String
    Dim cas_doublon_TECID As Long, cas_date_invalide As Long, cas_doublon_prof As Long, cas_doublon_client As Long
    Dim cas_date_fact_invalide As Long, cas_date_facture_future As Long, cas_date_future As Long
    Dim cas_hres_invalide As Long, cas_estFacturable_invalide As Long, cas_estFacturee_invalide As Long
    Dim cas_estDetruit_invalide As Long
    Dim total_hres_inscrites As Double, total_hres_detruites As Double, total_hres_facturees As Double
    Dim total_hres_facturable As Double, total_hres_TEC As Double, total_hres_non_facturable As Double
    Dim keyDate As String
    
    minDate = "12/31/2999"
    
    Dim isTECValid As Boolean
    isTECValid = True
    
    Dim arrHres(1 To 10000, 1 To 6) As Variant
    Dim arrRow As Integer, pArr As Integer, rArr As Integer
    
    'Sommaire par Date de charge (validation du format de date)
    Dim dictDateCharge As Object
    Set dictDateCharge = CreateObject("Scripting.Dictionary")
    Dim yy As Integer, mm As Integer, dd As Integer
    
    'Sommaire par TimeStamp (validation du format de date)
    Dim dictTimeStamp As Object
    Set dictTimeStamp = CreateObject("Scripting.Dictionary")
    
    Dim strDict As String

    'Lecture et analyse des TEC (TEC_Local)
    For i = LBound(arrTEC_Local_Data, 1) To UBound(arrTEC_Local_Data, 1)
        tecID = arrTEC_Local_Data(i, 1)
        If tecID > maxTECID Then
            maxTECID = tecID
        End If
        'ProfessionnelID
        profID = arrTEC_Local_Data(i, 2)
        'Professionnel
        prof = arrTEC_Local_Data(i, 3)
        'Date
        dateTEC = arrTEC_Local_Data(i, 4)
        testDate = IsDate(dateTEC)
        If testDate = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID =" & tecID & " a une date INVALIDE '" & dateTEC & " !!!")
            r = r + 1
            isTECValid = False
            cas_date_invalide = cas_date_invalide + 1
        Else
            If dateTEC < minDate Then minDate = dateTEC
            If dateTEC > maxDate Then maxDate = dateTEC
        End If
        If dateTEC > Date Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID =" & tecID & " a une date FUTURE '" & dateTEC & " !!!")
            r = r + 1
            isTECValid = False
            cas_date_future = cas_date_future + 1
        End If
        If dateTEC <> Int(dateTEC) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La date du TEC '" & dateTEC & "' n'est pas du bon format (H:M:S) pour le TECID =" & tecID)
            r = r + 1
            isTECValid = False
        End If
        
        'Validate clientCode
        codeClient = Trim(arrTEC_Local_Data(i, 5))
        If dictClient.Exists(codeClient) = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le code de client '" & codeClient & "' est INVALIDE !!!")
            r = r + 1
            isTECValid = False
        End If
        nomCLient = arrTEC_Local_Data(i, 6)
        hres = arrTEC_Local_Data(i, 8)
        testHres = IsNumeric(hres)
        If testHres = False Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            r = r + 1
            isTECValid = False
            cas_hres_invalide = cas_hres_invalide + 1
        End If
        estFacturable = arrTEC_Local_Data(i, 10)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            r = r + 1
            isTECValid = False
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If

        'Analyse de la date de charge et du TimeStamp pour les dernières entrées
        If arrTEC_Local_Data(i, 1) > lastTECIDReported And arrTEC_Local_Data(i, fTECEstDetruit) = "FAUX" Then
            'Date de la charge
            yy = year(arrTEC_Local_Data(i, 4))
            mm = month(arrTEC_Local_Data(i, 4))
            dd = day(arrTEC_Local_Data(i, 4))
            If yy = 2024 And mm = 12 And dd = 27 Then Stop
            strDict = Format$(DateSerial(yy, mm, dd), "yyyy-mm-dd") & " - " & _
                                Fn_Pad_A_String(CStr(arrTEC_Local_Data(i, 3)), " ", 5, "R")
            If dictDateCharge.Exists(strDict) Then
                dictDateCharge(strDict) = dictDateCharge(strDict) + arrTEC_Local_Data(i, 8)
            Else
                dictDateCharge.Add strDict, arrTEC_Local_Data(i, 8)
            End If
            'TimeStamp
            yy = year(arrTEC_Local_Data(i, 11))
            mm = month(arrTEC_Local_Data(i, 11))
            dd = day(arrTEC_Local_Data(i, 11))
            strDict = Format$(DateSerial(yy, mm, dd), "yyyy-mm-dd") & " - " & _
                                Fn_Pad_A_String(CStr(arrTEC_Local_Data(i, 3)), " ", 5, "R")
            If dictTimeStamp.Exists(strDict) Then
                dictTimeStamp(strDict) = dictTimeStamp(strDict) + 1
            Else
                dictTimeStamp.Add strDict, 1
            End If
        End If

        estFacturee = UCase(arrTEC_Local_Data(i, 12))
        If InStr("Vrai^VRAI^Faux^FAUX^", estFacturee & "^") = 0 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            r = r + 1
            isTECValid = False
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        
        If arrTEC_Local_Data(i, 13) <> "" Then
            dateFact = arrTEC_Local_Data(i, 13)
            testDate = IsDate(dateFact)
            If testDate = False Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID =" & tecID & " a une date de facture INVALIDE '" & dateFact & " !!!")
                r = r + 1
                isTECValid = False
                cas_date_fact_invalide = cas_date_fact_invalide + 1
            End If
            If dateFact > Now() Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID =" & tecID & " a une date de facture FUTURE '" & dateFact & " !!!")
                r = r + 1
                isTECValid = False
                cas_date_facture_future = cas_date_facture_future + 1
            End If
            If dateFact <> Int(dateFact) Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** La date de la facture '" & dateFact & "' n'est pas du bon format (H:M:S) pour le TECID =" & tecID)
                r = r + 1
                isTECValid = False
            End If
        End If
        
        estDetruit = arrTEC_Local_Data(i, 14)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            r = r + 1
            isTECValid = False
            cas_estDetruit_invalide = cas_estDetruit_invalide + 1
        End If
        
        invNo = CStr(arrTEC_Local_Data(i, 16))
        If Len(invNo) > 0 Then
            If estFacturee <> "VRAI" Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & _
                    " - Incongruité entre le numéro de facture '" & invNo & "' et " & _
                    "'estFacture' qui vaut '" & estFacturee & "'")
                r = r + 1
                isTECValid = False
            End If
            If Not invNo = "Radiation" Then
                If dictFacture.Exists(invNo) = False Then
                    Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & _
                        " - Le numéro de facture '" & invNo & "' " & _
                        "n'existe pas dans le fichier FAC_Entête")
                    r = r + 1
                    isTECValid = False
                Else 'Accumule les heures pour cette facture
                    If dictFactureHres.Exists(invNo) = True Then
                        dictFactureHres(invNo) = dictFactureHres(invNo) + arrTEC_Local_Data(i, 8)
                    Else
                        dictFactureHres.Add invNo, arrTEC_Local_Data(i, 8)
                    End If
                End If
            End If
        Else
            If estFacturee = "Vrai" Or estFacturee = "VRAI" Then
                Call AddMessageToWorkSheet(wsOutput, r, 2, "********** TECID = " & tecID & _
                    " - Incongruité entre le numéro de facture vide et " & _
                    "'estFacture' qui vaut '" & estFacturee & "'")
                r = r + 1
                isTECValid = False
            End If
        End If

        'Accumule les heures
        Dim h(1 To 6) As Double
        
        'Heures INSCRITES
        total_hres_inscrites = total_hres_inscrites + hres
        h(1) = hres
        
        'Heures DÉTRUITES
        h(2) = 0
        If estDetruit = "Vrai" Then
            total_hres_detruites = total_hres_detruites + hres
            h(2) = hres
            hres = 0 'Il ne reste plus d'heures...
        End If
        
        'Heures FACTURABLES
        h(3) = 0
        If hres <> 0 And estFacturable = "Vrai" And Fn_Is_Client_Facturable(codeClient) = True Then
                total_hres_facturable = total_hres_facturable + hres
                h(3) = hres
        End If
        
        'Heures NON-FACTURABLES
        h(4) = 0
        If hres <> 0 Then
            total_hres_non_facturable = total_hres_non_facturable + hres - h(3)
            h(4) = hres - h(3)
        End If
        
        'Heures FACTURÉES
        h(5) = 0
        If estFacturee = "Vrai" And Fn_Is_Client_Facturable(codeClient) = True Then
                total_hres_facturees = total_hres_facturees + hres
                h(5) = hres
        End If
        
        'Heures TEC = Heures Facturables - Heures facturées
        If h(3) Then
            h(6) = h(3) - h(5)
        Else
            h(6) = 0
        End If
        
        If h(1) - h(2) <> h(3) + h(4) Then
            Debug.Print "#020 - " & i & " Écart - " & tecID & " " & prof & " " & dateTEC & " " & h(1) & " " & h(2) & " vs. " & h(3) & " " & h(4)
            Stop
        End If
        
        'Dictionaries
        If dict_TECID.Exists(tecID) = False Then
            dict_TECID.Add tecID, 0
        Else
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Le TECID '" & tecID & "' est un doublon pour la ligne '" & i & "'")
            r = r + 1
            isTECValid = False
            cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        If dict_prof.Exists(prof & "-" & profID) = False Then
            dict_prof.Add prof & "-" & profID, 0
        End If
    Next i
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arrTEC_Local_Data, 1) - HeaderRow, "##,##0") & " charges de temps ont été analysées!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arrTEC_Local_Data, 1) - HeaderRow
    
    If cas_doublon_TECID = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucun doublon de TECID")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_doublon_TECID & " cas de doublons pour les TECID")
        r = r + 1
    End If
    
    If cas_date_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune date INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        r = r + 1
    End If
    
    If cas_date_future = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune date dans le futur")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_date_future & " cas de date FUTURE")
        r = r + 1
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    r = r + 1
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    r = r + 1
    
    If cas_hres_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune heures INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturable_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune valeur 'estFacturable' n'est INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturee_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune valeur 'estFacturee' n'est INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        r = r + 1
    End If
    
    If cas_date_fact_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune date de facture INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_date_fact_invalide & " cas de date de facture INVALIDE")
        r = r + 1
    End If
    
    If cas_estDetruit_invalide = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Aucune valeur 'estDetruit' n'est INVALIDE")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        r = r + 1
    End If
    
    Call AddMessageToWorkSheet(wsOutput, r, 2, "Vérification des Heures Facturées par Facture")
    r = r + 1
    
    'Vérification des heures facturées selon 2 sources (TEC_Local vs. FAC_Détails)
    Dim key As Variant
    Dim totalHoursBilled As Double
    Dim cas_Heures_Differentes As Integer
    
    For Each key In dictFacture.keys
        totalHoursBilled = Fn_Get_TEC_Total_Invoice_AF(CStr(key), "Heures")
        If Round(totalHoursBilled, 2) <> Round(dictFactureHres(key), 2) Then
            Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Facture '" & CStr(key) & _
                    "', il y a un écart d'heures facturées entre TEC_Local & FAC_Détails - " & _
                        Round(dictFactureHres(key), 2) & " vs. " & Round(totalHoursBilled, 2))
            r = r + 1
            cas_Heures_Differentes = cas_Heures_Differentes + 1
        End If
    Next key

    If cas_Heures_Differentes = 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "       Toutes les heures facturées balancent, selon les 2 sources")
        r = r + 1
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "********** Certaines factures sont à vérifier pour que les heures facturées balancent, selon les 2 sources")
        r = r + 1
        isTECValid = False
    End If
        
    Call AddMessageToWorkSheet(wsOutput, r, 2, "La somme des heures SAISIES donne ces résultats:")
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Heures SAISIES         : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Heures détruites       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Heures NETTES          : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "              Non_facturables : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "              Facturables     : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call AddMessageToWorkSheet(wsOutput, r, 2, "       Heures facturées       : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable - total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.Value = "       Heures TEC             : " & formattedHours
    rng.Characters(InStr(rng.Value, ":") + 2, Len(formattedHours)).Font.COLOR = vbRed
    rng.Characters(InStr(rng.Value, ":") + 2, Len(formattedHours)).Font.Bold = True
    r = r + 1
    
    Dim keys() As Variant
    
    'Tri & impression de dictDateCharge
    If dictDateCharge.count > 0 Then
        'Un peu de couleur
        Set rng = wsOutput.Range("B" & r)
        rng.Value = "Sommaire des heures selon la DATE de la charge (" & maxTECID & ")"
        rng.Characters(InStr(rng.Value, "(") + 1, Len(maxTECID)).Font.COLOR = vbGreen
        rng.Characters(InStr(rng.Value, "(") + 1, Len(maxTECID)).Font.Bold = True
        r = r + 1
    
        keys = dictDateCharge.keys
        Call Fn_Quick_Sort(keys, LBound(keys), UBound(keys))
        'Parcourir les clés triées et afficher les heures
        For i = LBound(keys) To UBound(keys)
            key = keys(i)
            formattedHours = Format$(dictDateCharge(key), "#0.00")
            formattedHours = String(6 - Len(formattedHours), " ") & formattedHours
            Call AddMessageToWorkSheet(wsOutput, r, 2, "       " & key & ":" & formattedHours & " heures")
            r = r + 1
        Next i
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "Aucune nouvelle saisie d'heures (TECID > " & lastTECIDReported & ") ")
        r = r + 1
    End If
    
    'Tri & impression de dictTimeStamp
    If dictTimeStamp.count > 0 Then
        Call AddMessageToWorkSheet(wsOutput, r, 2, "Sommaire des heures saisies selon le 'TIMESTAMP'")
        r = r + 1
        keys = dictTimeStamp.keys
        Call Fn_Quick_Sort(keys, LBound(keys), UBound(keys))
        'Parcourir les clés triées et afficher les valeurs
        For i = LBound(keys) To UBound(keys)
            key = keys(i)
            formattedHours = Format$(dictTimeStamp(key), "##0")
            formattedHours = String(6 - Len(formattedHours), " ") & formattedHours
            Call AddMessageToWorkSheet(wsOutput, r, 2, "       " & key & ":" & formattedHours & " entrée(s)")
            r = r + 1
        Next i
    Else
        Call AddMessageToWorkSheet(wsOutput, r, 2, "Aucune nouvelle saisie d'heures (TECID > " & lastTECIDReported & ") ")
        r = r + 1
    End If
    r = r + 1
    
    'Cas problème dans cette vérification ?
    If isTECValid = False Then
        verificationIntegriteOK = False
    End If

Clean_Exit:

    'Libérer la mémoire
    Set dictDateCharge = Nothing
    Set dictFacture = Nothing
    Set dictTimeStamp = Nothing
    Set dict_prof = Nothing
    Set dict_TECID = Nothing
    Set key = Nothing
    Set rng = Nothing
    Set rngTEC_LocalData = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:checkTEC", startTime)

End Sub

Sub Make_It_As_Header(r As Range)

    With r
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .COLOR = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .size = 9
            .Italic = True
            .Bold = True
        End With
        .HorizontalAlignment = xlCenter
    End With
    
    Dim wsName As String
    wsName = r.Worksheet.Name
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(wsName)
    ws.Columns.AutoFit
    
    'Libérer la mémoire
    Set r = Nothing
    Set ws = Nothing

End Sub

Sub AddMessageToWorkSheet(ws As Worksheet, r As Long, c As Long, m As String)

    ws.Cells(r, c).Value = m
    If c = 1 Then
        ws.Cells(r, c).Font.Bold = True
    End If

End Sub

Sub ApplyConditionalFormatting(rng As Range, headerRows As Long, Optional EmptyLine As Boolean = False)

    'Avons-nous un Range valide ?
    If rng Is Nothing Or rng.Rows.count <= headerRows Then
        Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = rng.Worksheet
    Dim dataRange As Range
    
   ' Définir la plage de données à laquelle appliquer la mise en forme conditionnelle, en
    'excluant les lignes d'en-tête
    Set dataRange = rng.Resize(rng.Rows.count - headerRows).offset(headerRows, 0)
    
    'Effacer les formats conditionnels existants sur la plage de données
    dataRange.Interior.ColorIndex = xlNone

    'Appliquer les couleurs en alternance
    Dim i As Long
    For i = 1 To dataRange.Rows.count
        'Vérifier la position réelle de la ligne dans la feuille
        If (dataRange.Rows(i).row + headerRows) Mod 2 = 0 Then
            dataRange.Rows(i).Interior.COLOR = RGB(173, 216, 230) ' Bleu pâle
        End If
    Next i
    
    'Libérer la mémoire
    Set dataRange = Nothing
    Set ws = Nothing
    
End Sub

Sub ApplyWorksheetFormat(ws As Worksheet, rng As Range, HeaderRow As Long)

    'Common stuff to all worksheets
    rng.EntireColumn.AutoFit 'Autofit all columns
    
    'Conditional Formatting (many steps)
    '1) Remove existing conditional formatting
        rng.Cells.FormatConditions.Delete 'Remove the worksheet conditional formatting
    
    '2) Define the usedRange to data only (exclude header row(s))
        Dim numRows As Long
        numRows = rng.CurrentRegion.Rows.count - HeaderRow
        Dim usedRange As Range
        If numRows > 0 Then
            On Error Resume Next
            Set usedRange = rng.offset(HeaderRow, 0).Resize(numRows, rng.Columns.count)
            On Error GoTo 0
        End If
    
'    '3) Add the standard conditional formatting
'        If Not usedRange Is Nothing Then
'            With usedRange
'                .FormatConditions.add Type:=xlExpression, _
'                    Formula1:="=MOD(LIGNE();2)=1"
'                .FormatConditions(.FormatConditions.count).SetFirstPriority
'                With .FormatConditions(1).Font
'                    .Strikethrough = False
'                    .TintAndShade = 0
'                End With
'                With .FormatConditions(1).Interior
'                    .PatternColorIndex = xlAutomatic
'                    .ThemeColor = xlThemeColorAccent1
'                    .TintAndShade = 0.799981688894314
'                End With
'                .FormatConditions(1).StopIfTrue = False
'            End With
'        Else
'            MsgBox "usedRange is Nothing!"
'        End If
        
    'Specific formats to worksheets
    Dim lastUsedRow As Long
    lastUsedRow = rng.Rows.count
    If lastUsedRow = HeaderRow Then
        Exit Sub
    End If
    
    Dim firstDataRow As Long
    firstDataRow = HeaderRow + 1
    
    Select Case rng.Worksheet.CodeName
        Case "wshBD_Clients"
            
        Case "wshBD_Fournisseurs"
            
        Case "wshDEB_Récurrent"
            With wshDEB_Récurrent
                .Range("A2:M" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("C2:C" & lastUsedRow & _
                     ", D2:D" & lastUsedRow & _
                     ", E2:E" & lastUsedRow & _
                     ", G2:G" & lastUsedRow).HorizontalAlignment = xlLeft
                With .Range("I2:M" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
                End With
                .Range("N2:N" & lastUsedRow).NumberFormat = "yyyy-mm-dd hh:mm:ss"
            End With
       
        Case "wshDEB_Trans"
            With wshDEB_Trans
                .Range("A2:R" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("C2:C" & lastUsedRow & ", " & _
                       "D2:D" & lastUsedRow & ", " & _
                       "F2:F" & lastUsedRow & ", " & _
                       "G2:G" & lastUsedRow & ", " & _
                       "I2:I" & lastUsedRow & ", " & _
                       "Q2:Q" & lastUsedRow).HorizontalAlignment = xlLeft
                With .Range("K2:P" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
                End With
                .Range("R2:R" & lastUsedRow).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                
                .Range("A1").CurrentRegion.EntireColumn.AutoFit
            End With
        
        Case "wshENC_Détails"
            With wshENC_Détails
                .Range("A2:A" & lastUsedRow & ", B2:B" & lastUsedRow & ", D2:D" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("C2:C" & lastUsedRow & ", E2:EB" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("E2:E" & lastUsedRow).HorizontalAlignment = xlRight
                
                .Range("A2:A" & lastUsedRow).NumberFormat = "0"
                .Range("D2:D" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("E2:E" & lastUsedRow).NumberFormat = "#,##0.00"
            End With
        
        Case "wshENC_Entête"
            With wshENC_Entête
                .Range("A2:A" & lastUsedRow & ", B2:B" & lastUsedRow & ", D2:D" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("C2:C" & lastUsedRow & ", E2:E" & lastUsedRow & ", G2:G" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("F2:F" & lastUsedRow).HorizontalAlignment = xlRight
                
                .Range("A2:A" & lastUsedRow).NumberFormat = "0"
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("F2:F" & lastUsedRow).NumberFormat = "#,##0.00 $"
            End With
        
        Case "wshFAC_Comptes_Clients"
            With wshFAC_Comptes_Clients
                .Range(.Cells(3, fFacCCInvNo), .Cells(lastUsedRow, fFacCCInvoiceDate)).HorizontalAlignment = xlCenter
                .Range(.Cells(3, fFacCCCodeClient), .Cells(lastUsedRow, fFacCCDueDate)).HorizontalAlignment = xlCenter
                .Range(.Cells(3, fFacCCCustomer), .Cells(lastUsedRow, fFacCCCustomer)).HorizontalAlignment = xlLeft
                .Range(.Cells(3, fFacCCTotal), .Cells(lastUsedRow, fFacCCBalance)).HorizontalAlignment = xlRight
                .Range(.Cells(3, fFacCCInvoiceDate), .Cells(lastUsedRow, fFacCCInvoiceDate)).NumberFormat = "yyyy-mm-dd"
                .Range(.Cells(3, fFacCCDueDate), .Cells(lastUsedRow, fFacCCDueDate)).NumberFormat = "yyyy-mm-dd"
                .Range(.Cells(3, fFacCCTotal), .Cells(lastUsedRow, fFacCCBalance)).NumberFormat = "###,##0.00 $"
                
                .Range("A1").CurrentRegion.EntireColumn.AutoFit
            End With
        
        Case "wshFAC_Détails"
            With usedRange
                .Range("A2:A" & lastUsedRow & ", C2:C" & lastUsedRow & ", F2:F" & lastUsedRow & ", G2:G" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("D2:E" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("C2:C" & lastUsedRow).NumberFormat = "#,##0.00"
                .Range("D2:E" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("H2:H" & lastUsedRow & ", J2:J" & lastUsedRow & ", L2:L" & lastUsedRow & ", N2:T" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("O2:O" & lastUsedRow & ", Q2:Q" & lastUsedRow).NumberFormat = "#0.000 %"
            End With
        
        Case "wshFAC_Entête"
            With wshFAC_Entête
                .Range("A2:D" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("E2:I" & lastUsedRow & ", K2:K" & lastUsedRow & ", M2:M" & lastUsedRow & ", O2:O" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("J2:J" & lastUsedRow & ", L2:L" & lastUsedRow & ", N2:N" & lastUsedRow & ", P2:V" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("J2:J" & lastUsedRow & ", L2:L" & lastUsedRow & ", N2:N" & lastUsedRow & ", P2:V" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("Q2:Q" & lastUsedRow & ",S2:S" & lastUsedRow).NumberFormat = "#0.000 %"
            End With

        Case "wshFAC_Projets_Détails"
            With wshFAC_Projets_Détails
                .Range("A2:A" & lastUsedRow & ", C2:G" & lastUsedRow & ", I2:J" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("F2:F" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("H2:I" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("H2:H" & lastUsedRow).NumberFormat = "#,##0.00"
                .Range("I2:I" & lastUsedRow).HorizontalAlignment = xlCenter
            End With
        
        Case "wshFAC_Projets_Entête"
            With wshFAC_Projets_Entête
                .Range("A2:A" & lastUsedRow & ", C2:D" & lastUsedRow & ", F2:F" & lastUsedRow & _
                       ", J2:J" & lastUsedRow & ", N2:N" & lastUsedRow & ", R2:R" & lastUsedRow & _
                       ", V2:V" & lastUsedRow & ", Z2:AA" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("E2:E" & lastUsedRow & ", I2:I" & lastUsedRow & ", M2:M" & lastUsedRow & _
                        ", Q2:Q" & lastUsedRow & ", U2:U" & lastUsedRow & ", Y2:Y" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("G2:H" & lastUsedRow).NumberFormat = "#,##0.00"
            End With
        
        Case "wshGL_EJ_Recurrente"
            With wshGL_EJ_Recurrente
                Union(.Range("A2:A" & lastUsedRow), _
                      .Range("C2:C" & lastUsedRow), _
                      .Range("H2:H" & lastUsedRow), _
                      .Range("K2:K" & lastUsedRow)).HorizontalAlignment = xlCenter
                Union(.Range("B2:B" & lastUsedRow), _
                      .Range("D2:D" & lastUsedRow), _
                      .Range("G2:G" & lastUsedRow), _
                      .Range("J2:J" & lastUsedRow)).HorizontalAlignment = xlLeft
                With .Range("E2:F" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "###,##0.00 $"
                End With
            End With
        
        Case "wshGL_Trans"
            With wshGL_Trans
                .Range("A2:J" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("C2:C" & lastUsedRow & _
                    ", D2:D" & lastUsedRow & _
                    ", F2:F" & lastUsedRow & _
                    ", I2:I" & lastUsedRow) _
                        .HorizontalAlignment = xlLeft
                With .Range("G2:H" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
                End With
                .Range("J2:J" & lastUsedRow).NumberFormat = "yyyy-mm-dd hh:mm:ss"
            End With
        
        Case "wshTEC_Local"
            With wshTEC_Local
                .Range("A2:P" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("F2:F" & lastUsedRow & ", G2:G" & lastUsedRow & ", I2:I" & lastUsedRow & _
                            ", O2:O" & lastUsedRow).HorizontalAlignment = xlLeft
                            
                .Range("H2:H" & lastUsedRow).NumberFormat = "#0.00"
                .Range("D2:D" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("K2:K" & lastUsedRow).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                .Columns("F").ColumnWidth = 40
                .Columns("G").ColumnWidth = 55
                .Columns("I").ColumnWidth = 20
            End With

    End Select

    'Libérer la mémoire
    Set usedRange = Nothing

End Sub

Sub Fix_Font_Size_And_Family(r As Range, ff As String, fs As Long)

    'r is the range
    'ff is the Font Family
    'fs is the Font Size
    
    With r.Font
        .Name = ff
        .size = fs
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With

End Sub

Sub Get_Deplacements_From_TEC()  '2024-09-05 @ 10:22

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli_Utils:Get_Deplacements_From_TEC", 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Mise en place de la feuille de sortie (output)
    Dim strOutput As String
    strOutput = "X_TEC_Déplacements"
    Call CreateOrReplaceWorksheet(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").Value = "Date"
    wsOutput.Range("B1").Value = "Date"
    wsOutput.Range("C1").Value = "Nom du client"
    wsOutput.Range("D1").Value = "Heures"
    wsOutput.Range("E1").Value = "Adresse_1"
    wsOutput.Range("F1").Value = "Adresse_2"
    wsOutput.Range("G1").Value = "Ville"
    wsOutput.Range("H1").Value = "Province"
    wsOutput.Range("I1").Value = "CodePostal"
    wsOutput.Range("J1").Value = "DistanceKM"
    wsOutput.Range("K1").Value = "Montant"
    Call Make_It_As_Header(wsOutput.Range("A1:K1"))
    
    'Feuille pour les clients
    Dim wsMF As Worksheet: Set wsMF = wshBD_Clients
    Dim lastUsedRowClientMF As Long
    lastUsedRowClientMF = wsMF.Cells(wsMF.Rows.count, 1).End(xlUp).row
    Dim rngClientsMF As Range
    Set rngClientsMF = wsMF.Range("A1:A" & lastUsedRowClientMF)
    
    'Get From and To Dates
    Dim dateFrom As Date, dateTo As Date
    dateFrom = wshAdmin.Range("MoisPrecDe").Value
    dateTo = wshAdmin.Range("MoisPrecA").Value
    
    'Analyse de TEC_Local
    Call TEC_Import_All
    
    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    Dim lastUsedRowTEC As Long
    lastUsedRowTEC = wsTEC.Cells(wsTEC.Rows.count, 1).End(xlUp).row
    Dim arr() As Variant
    
    'Copier le range en mémoire
    Call Tx_Range_2_2D_Array(wsTEC.Range("A1:P" & lastUsedRowTEC), arr, 2)
    
    'Mise en place d'un tableau pour recevoir les résultats (performance)
    Dim output() As Variant
    ReDim output(1 To UBound(arr, 1), 1 To UBound(arr, 2))
    Dim rowOutput As Long
    rowOutput = 1
    
    Dim clientData As Variant
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 3) = "GC" And UCase(arr(i, 14)) <> "VRAI" Then
            If arr(i, 4) >= CLng(dateFrom) And arr(i, 4) <= CLng(dateTo) Then
                output(rowOutput, 1) = arr(i, 4)
                output(rowOutput, 2) = arr(i, 4)
                output(rowOutput, 4) = arr(i, 8)
                clientData = Fn_Rechercher_Client_Par_ID(Trim(arr(i, 5)), wsMF)
                If IsArray(clientData) Then
                    output(rowOutput, 3) = clientData(1, fClntFMClientNom)
                    output(rowOutput, 5) = clientData(1, fClntFMAdresse1)
                    output(rowOutput, 6) = clientData(1, fClntFMAdresse2)
                    output(rowOutput, 7) = clientData(1, fClntFMVille)
                    output(rowOutput, 8) = clientData(1, fClntFMProvince)
                    output(rowOutput, 9) = clientData(1, fClntFMCodePostal)
                End If
                rowOutput = rowOutput + 1
            End If
        End If
    Next i
    
    'Copier le tableau dans le range
    Call Tx_2D_Array_2_Range(output, wsOutput.Range("A2:I" & UBound(output, 1)), True, 1)
    
    'Tri des données
    With wsOutput.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsOutput.Range("B2"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortTextAsNumbers 'Sort Date
        .SortFields.Add key:=wshTEC_Local.Range("C2"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort on Client's name
        .SortFields.Add key:=wshTEC_Local.Range("D2"), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal 'Sort on Hours
        .SetRange wsOutput.Range("A2:K" & rowOutput - 1) 'Set Range
        .Apply 'Apply Sort
     End With
    
    'Ajustement des formats
    With wsOutput
        .Range("A2:B" & rowOutput + 1).NumberFormat = wshAdmin.Range("B1").Value
        .Range("D2:D" & rowOutput + 1).NumberFormat = "##0.00"
        .Range("A2:K" & rowOutput + 1).Font.Name = "Aptos Narrow"
        .Range("A2:K" & rowOutput + 1).Font.size = 10
        .Columns.AutoFit
    End With
    
    'Améliore le Look (saute 1 ligne entre chaque jour)
    For i = rowOutput To 3 Step -1
        If Len(Trim(wsOutput.Cells(i, 3).Value)) > 0 Then
            If wsOutput.Cells(i, 2).Value <> wsOutput.Cells(i - 1, 2).Value Then
                wsOutput.Rows(i).Insert Shift:=xlDown
                wsOutput.Cells(i, 1).Value = wsOutput.Cells(i - 1, 2).Value
            End If
        End If
    Next i
    
    rowOutput = wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).row
    
    'Améliore le Look (cache la date, le client et l'adresse si deux charges & +)
    Dim base As String
    For i = 2 To rowOutput
        If i = 2 Then
            base = wsOutput.Cells(i, 2).Value & wsOutput.Cells(i, 3).Value
        End If
        If i > 2 And Len(wsOutput.Cells(i, 2).Value) > 0 Then
            If wsOutput.Cells(i, 2).Value & wsOutput.Cells(i, 3).Value = base Then
                wsOutput.Cells(i, 2).Value = ""
                wsOutput.Cells(i, 3).Value = ""
                wsOutput.Cells(i, 5).Value = ""
                wsOutput.Cells(i, 6).Value = ""
                wsOutput.Cells(i, 7).Value = ""
                wsOutput.Cells(i, 8).Value = ""
                wsOutput.Cells(i, 9).Value = ""
            Else
                base = wsOutput.Cells(i, 2).Value & wsOutput.Cells(i, 3).Value
            End If
        End If
    Next i
    
    'Result print setup - 2024-08-05 @ 05:16
    rowOutput = wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).row
    
    For i = 3 To rowOutput
        If wsOutput.Cells(i, 1).Value > wsOutput.Cells(i - 1, 1).Value Then
            wsOutput.Cells(i, 2).Font.Bold = True
        Else
            wsOutput.Cells(i, 2).Value = ""
        End If
    Next i
    
    'Première date est en caractère gras
    wsOutput.Cells(2, 2).Font.Bold = True
    rowOutput = rowOutput + 2
    wsOutput.Range("A" & rowOutput).Value = "**** " & Format$(lastUsedRowTEC - 2, "###,##0") & _
                                        " charges de temps analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("B2:K" & rowOutput)
    Call ApplyConditionalFormatting(rngArea, 1, True)

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Setup print parameters
'    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:I" & rowOutput)
    Dim header1 As String: header1 = "Liste des TEC pour Guillaume"
    Dim header2 As String: header2 = "Période du " & dateFrom & " au " & dateTo
    Call Simple_Print_Setup(wsOutput, rngArea, header1, header2, "$1:$1", "P")
    
    'Libérer la mémoire
    Set rngArea = Nothing
    Set rngClientsMF = Nothing
    Set wsOutput = Nothing
    Set wsMF = Nothing
    Set wsTEC = Nothing
    
    Call Log_Record("modAppli_Utils:Get_Deplacements_From_TEC", startTime)

End Sub

Sub Get_Deplacements_From_TEC_OK()  '2024-09-05 @ 10:22

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli_Utils:Get_Deplacements_From_TEC", 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Mise en place de la feuille de sortie (output)
    Dim strOutput As String
    strOutput = "X_TEC_Déplacements"
    Call CreateOrReplaceWorksheet(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").Value = "Date"
    wsOutput.Range("B1").Value = "Date"
    wsOutput.Range("C1").Value = "Nom du client"
    wsOutput.Range("D1").Value = "Heures"
    wsOutput.Range("E1").Value = "Adresse_1"
    wsOutput.Range("F1").Value = "Adresse_2"
    wsOutput.Range("G1").Value = "Ville"
    wsOutput.Range("H1").Value = "Province"
    wsOutput.Range("I1").Value = "CodePostal"
    wsOutput.Range("J1").Value = "DistanceKM"
    wsOutput.Range("K1").Value = "Montant"
    Call Make_It_As_Header(wsOutput.Range("A1:K1"))
    
    'Feuille pour les clients
    Dim wsMF As Worksheet: Set wsMF = wshBD_Clients
    Dim lastUsedRowClientMF As Long
    lastUsedRowClientMF = wsMF.Cells(wsMF.Rows.count, 1).End(xlUp).row
    Dim rngClientsMF As Range
    Set rngClientsMF = wsMF.Range("A1:A" & lastUsedRowClientMF)
    
    'Get From and To Dates
    Dim dateFrom As Date, dateTo As Date
    dateFrom = wshAdmin.Range("MoisPrecDe").Value
    dateTo = wshAdmin.Range("MoisPrecA").Value
    
    'Analyse de TEC_Local
    Call TEC_Import_All
    
    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    Dim lastUsedRowTEC As Long
    lastUsedRowTEC = wsTEC.Cells(wsTEC.Rows.count, 1).End(xlUp).row
    Dim arr() As Variant
    
    Call Tx_Range_2_2D_Array(wsTEC.Range("A1:P" & lastUsedRowTEC), arr, 2)
    
    Dim rowOutput As Long
    rowOutput = 2 'Skip the header
    Dim clientData As Variant
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 3) = "GC" And _
            arr(i, 4) >= dateFrom And _
            arr(i, 4) <= dateTo And _
            UCase(arr(i, 14)) <> "VRAI" Then
                wsOutput.Cells(rowOutput, 1).Value = arr(i, 4)
                wsOutput.Cells(rowOutput, 1).NumberFormat = wshAdmin.Range("B1").Value
                wsOutput.Cells(rowOutput, 2).Value = arr(i, 4)
                wsOutput.Cells(rowOutput, 2).NumberFormat = wshAdmin.Range("B1").Value
                wsOutput.Cells(rowOutput, 4).Value = arr(i, 8)
                clientData = Fn_Rechercher_Client_Par_ID(Trim(arr(i, 5)), wsMF)
                If IsArray(clientData) Then
                    wsOutput.Cells(rowOutput, 3).Value = clientData(1, fClntFMClientNom)
                    wsOutput.Cells(rowOutput, 5).Value = clientData(1, fClntFMAdresse1)
                    wsOutput.Cells(rowOutput, 6).Value = clientData(1, fClntFMAdresse2)
                    wsOutput.Cells(rowOutput, 7).Value = clientData(1, fClntFMVille)
                    wsOutput.Cells(rowOutput, 8).Value = clientData(1, fClntFMProvince)
                    wsOutput.Cells(rowOutput, 9).Value = clientData(1, fClntFMCodePostal)
                End If
                rowOutput = rowOutput + 1
        End If
    Next i
    
    'Colonne des Heures
    wsOutput.Range("D2:D" & rowOutput - 1).NumberFormat = "##0.00"
    
    'Tri des données
    With wsOutput.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsOutput.Range("B2"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortTextAsNumbers 'Sort Date
        .SortFields.Add key:=wshTEC_Local.Range("C2"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort on Client's name
        .SortFields.Add key:=wshTEC_Local.Range("D2"), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal 'Sort on Hours
        .SetRange wsOutput.Range("A2:K" & rowOutput - 1) 'Set Range
        .Apply 'Apply Sort
     End With
    
    wsOutput.Columns.AutoFit

    'Améliore le Look (saute 1 ligne entre chaque jour)
    For i = rowOutput To 3 Step -1
        If Len(Trim(wsOutput.Cells(i, 3).Value)) > 0 Then
            If wsOutput.Cells(i, 2).Value <> wsOutput.Cells(i - 1, 2).Value Then
                wsOutput.Rows(i).Insert Shift:=xlDown
                wsOutput.Cells(i, 1).Value = wsOutput.Cells(i - 1, 2).Value
            End If
        End If
    Next i
    
    rowOutput = wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).row
    
    'Améliore le Look (cache la date, le client et l'adresse si deux charges & +)
    Dim base As String
    For i = 2 To rowOutput
        If i = 2 Then
            base = wsOutput.Cells(i, 2).Value & wsOutput.Cells(i, 3).Value
        End If
        If i > 2 And Len(wsOutput.Cells(i, 2).Value) > 0 Then
            If wsOutput.Cells(i, 2).Value & wsOutput.Cells(i, 3).Value = base Then
                wsOutput.Cells(i, 2).Value = ""
                wsOutput.Cells(i, 3).Value = ""
                wsOutput.Cells(i, 5).Value = ""
                wsOutput.Cells(i, 6).Value = ""
                wsOutput.Cells(i, 7).Value = ""
                wsOutput.Cells(i, 8).Value = ""
                wsOutput.Cells(i, 9).Value = ""
            Else
                base = wsOutput.Cells(i, 2).Value & wsOutput.Cells(i, 3).Value
            End If
        End If
    Next i
    
    'Result print setup - 2024-08-05 @ 05:16
    rowOutput = wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).row
    
    For i = 3 To rowOutput
        If wsOutput.Cells(i, 1).Value > wsOutput.Cells(i - 1, 1).Value Then
            wsOutput.Cells(i, 2).Font.Bold = True
        Else
            wsOutput.Cells(i, 2).Value = ""
        End If
    Next i
    'Première date est en caractère gras
    wsOutput.Cells(2, 2).Font.Bold = True
    rowOutput = rowOutput + 2
    wsOutput.Range("A" & rowOutput).Value = "**** " & Format$(lastUsedRowTEC - 2, "###,##0") & _
                                        " charges de temps analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("B2:K" & rowOutput)
    Call ApplyConditionalFormatting(rngArea, 1, True)

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Setup print parameters
'    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:I" & rowOutput)
    Dim header1 As String: header1 = "Liste des TEC pour Guillaume"
    Dim header2 As String: header2 = "Période du " & dateFrom & " au " & dateTo
    Call Simple_Print_Setup(wsOutput, rngArea, header1, header2, "$1:$1", "P")
    
    'Libérer la mémoire
    Set rngArea = Nothing
    Set rngClientsMF = Nothing
    Set wsOutput = Nothing
    Set wsMF = Nothing
    Set wsTEC = Nothing
    
    Call Log_Record("modAppli_Utils:Get_Deplacements_From_TEC", startTime)

End Sub

Sub Get_Date_Derniere_Modification(fileName As String, ByRef ddm As Date, _
                                    ByRef jours As Long, ByRef heures As Long, _
                                    ByRef minutes As Long, ByRef secondes As Long)
    
    'Créer une instance de FileSystemObject
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'Obtenir le fichier
    Dim fichier As Object: Set fichier = FSO.GetFile(fileName)
    
    'Récupérer la date et l'heure de la dernière modification
    ddm = fichier.DateLastModified
    
    'Calculer la différence (jours) entre maintenant et la date de la dernière modification
    Dim diff As Double
    diff = Now - ddm
    
    'Convertir la différence en jours, heures, minutes et secondes
    jours = Int(diff)
    heures = Int((diff - jours) * 24)
    minutes = Int(((diff - jours) * 24 - heures) * 60)
    secondes = Int(((((diff - jours) * 24 - heures) * 60) - minutes) * 60)
    
    ' Libérer les objets
    Set fichier = Nothing
    Set FSO = Nothing
    
End Sub

Sub Dynamic_Range_Redefine_Plan_Comptable() '2024-07-04 @ 10:39
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli_Utils:Dynamic_Range_Redefine_Plan_Comptable", 0)

    'Redefine - dnrPlanComptable_Description_Only
    'Delete existing dynamic named range (assuming it could exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptable_Description_Only").Delete
    On Error GoTo 0
    
    'Define a new dynamic named range for 'dnrPlanComptable_Description_Only'
    Dim newRangeFormula As String
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,1)"
    
    'Create the new dynamic named range
    ThisWorkbook.Names.Add Name:="dnrPlanComptable_Description_Only", RefersTo:=newRangeFormula
    
    'Redefine - dnrPlanComptable_All
    'Delete existing dynamic named range (assuming it could exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptable_All").Delete
    On Error GoTo 0
    
    'Define a new dynamic named range for 'dnrPlanComptable_All'
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,4)"
    
    'Create the new dynamic named range
    ThisWorkbook.Names.Add Name:="dnrPlanComptable_All", RefersTo:=newRangeFormula
    
    Call Log_Record("modImport:Dynamic_Range_Redefine_Plan_Comptable", startTime)

End Sub

Sub Remplir_Plage_Avec_Couleur(ByVal plage As Range, ByVal couleurRVB As Long)

    If Not plage Is Nothing Then
        Dim cellule As Range
        'Parcourt toutes les cellules de la plage (contiguës ou non)
        For Each cellule In plage
            On Error Resume Next
            cellule.Interior.COLOR = couleurRVB
            On Error GoTo 0
        Next cellule
    Else
        MsgBox "La plage spécifiée est invalide.", vbExclamation, "Procédure 'Remplir_Plage_Avec_Couleur'"
    End If
    
End Sub

Sub Paint_A_Range(rng As Range, colorRGB As String)
   
    Dim cell As Variant
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With

End Sub

Sub ChargerRangeDansDictionnaire(ByRef dict As Object, ByVal rng As Range, Optional colValeurOffset As Long = 0)

    'Créer un dictionnaire si non initialisé
    If dict Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
    End If

    'Parcourir chaque cellule de la plage et ajouter au dictionnaire
    Dim cell As Range
    Dim clé As Variant
    Dim valeur As Variant
    For Each cell In rng
        clé = cell.Value
        valeur = cell.offset(0, colValeurOffset).Value 'Colonne adjacente ou selon décalage

        'Ajouter au dictionnaire si la clé n'existe pas déjà
        If Not dict.Exists(clé) Then
            dict.Add clé, valeur
        End If
    Next cell
    
End Sub

Sub ExempleUtilisation()

    Dim dict As Object
    
    'Définir la feuille de calcul et la plage
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion.offset(2, 0)
    'Redimensionner la plage pour inclure uniquement les lignes restantes
    Set rng = rng.Resize(rng.Rows.count - 2, rng.Columns.count)

    'Charger les données dans un dictionnaire
    Call ChargerRangeDansDictionnaire(dict, rng, 2) ' 2 = Décalage de colonne pour les valeurs (colonne C)

'    'Afficher le contenu du dictionnaire
'    Dim clé As Variant
'    For Each clé In dict.keys
'        Debug.Print "Clé: " & clé & ", Valeur: " & dict(clé)
'    Next clé
'
    'Nettoyer la mémoire
    Set rng = Nothing
    Set ws = Nothing
    
End Sub

