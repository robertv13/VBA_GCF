Attribute VB_Name = "modAppli_Utils"
Option Explicit

Public Sub ConvertRangeBooleanToText(rng As Range)

    Dim cell As Range
    For Each cell In rng
        Select Case cell.value
            Case 0, "False" 'False
                cell.value = "FAUX"
            Case -1, "True" 'True
                cell.value = "VRAI"
            Case "VRAI", "FAUX"
                
            Case Else
                MsgBox cell.value & " est une valeur INVALIDE pour la cellule " & cell.Address & " de la feuille TEC_Local"
        End Select
    Next cell

    'Libérer la mémoire
    Set cell = Nothing
    
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
        
        .TopMargin = Application.InchesToPoints(0.8)
        .LeftMargin = Application.InchesToPoints(0.1)
        .RightMargin = Application.InchesToPoints(0.1)
        .BottomMargin = Application.InchesToPoints(0.5)
        
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
'    If Err.Number <> 0 Then
'        MsgBox "Error setting PrintCommunication to True: " & Err.Description, vbCritical
'    End If
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

Public Sub ArrayToRange(ByRef data As Variant _
                        , ByVal outRange As Range _
                        , Optional ByVal clearExistingData As Boolean = True _
                        , Optional ByVal clearExistingHeaderSize As Long = 1)
                        
    If clearExistingData = True Then
        outRange.CurrentRegion.Offset(clearExistingHeaderSize).ClearContents
    End If
    
    Dim rows As Long, columns As Long
    rows = UBound(data, 1) - LBound(data, 1) + 1
    columns = UBound(data, 2) - LBound(data, 2) + 1
    outRange.Resize(rows, columns).value = data
    
End Sub

Sub CreateOrReplaceWorksheet(wsName As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli_Utils:CreateOrReplaceWorksheet", 0)
    
    'Check if the worksheet exists
    Dim ws As Worksheet
    Dim wsExists As Boolean
    For Each ws In ThisWorkbook.Worksheets
        wsExists = False
        If ws.Name = wsName Then
            wsExists = True
            Exit For
        End If
    Next ws
    
    'If the worksheet exists, delete it
    If wsExists Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    'Add the new worksheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=wshMenu)
    ws.Name = wsName

    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modAppli_Utils:CreateOrReplaceWorksheet", startTime)

End Sub

Public Sub Integrity_Verification() '2024-07-06 @ 12:56

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:Integrity_Verification", 0)

    Application.ScreenUpdating = False
    
    Call Erase_And_Create_Worksheet("X_Analyse_Intégrité")
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    wsOutput.Unprotect
    wsOutput.Range("A1").value = "Feuille"
    wsOutput.Range("B1").value = "Message"
    wsOutput.Range("C1").value = "TimeStamp"
    wsOutput.columns("C").NumberFormat = wshAdmin.Range("B1").value & " hh:mm:ss"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"))

    Application.ScreenUpdating = True

    'Data starts at row 2
    Dim r As Long: r = 2
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Répertoire utilisé")
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, wshAdmin.Range("FolderSharedData").value & DATA_PATH)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1

    'Fichier utilisé
    Dim masterFileName As String
    masterFileName = "GCF_BD_MASTER.xlsx"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Fichier utilisé")
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, masterFileName)
    r = r + 1
    
    'Date dernière modification du fichier Maître
    Dim fullFileName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & DATA_PATH & Application.PathSeparator & masterFileName
    Dim ddm As Date
    Dim j As Long, h As Long, m As Long, s As Long
    Call Get_Date_Derniere_Modification(fullFileName, ddm, j, h, m, s)
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Date dern. modification")
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.value = Format$(ddm, wshAdmin.Range("B1").value & " hh:mm:ss") & _
            " soit " & j & " jours, " & h & " heures, " & m & " minutes et " & s & " secondes"
    rng.Characters(1, 19).Font.Color = vbRed
    rng.Characters(1, 19).Font.Bold = True

    r = r + 2
    
    Dim readRows As Long
    
    'dnrPlanComptable ----------------------------------------------------- Plan Comptable
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Plan Comptable")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    
    Call check_Plan_Comptable(r, readRows)
    
    'wshBD_Clients --------------------------------------------------------------- Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "BD_Clients")
    
    Call Client_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_Clients(r, readRows)
    
    'wshBD_Fournisseurs ----------------------------------------------------- Fournisseurs
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "BD_Fournisseurs")
    
    Call Fournisseur_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_Fournisseurs(r, readRows)
    
    'wshFAC_Entête ------------------------------------------------------------ FAC_Entête
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Entête")
    
    Call FAC_Entête_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Entête a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Entête(r, readRows)
    
    'wshFAC_Détails ---------------------------------------------------------- FAC_Détails
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Détails")
    
    Call FAC_Détails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Détails a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Détails(r, readRows)
    
    'wshFAC_Comptes_Clients ------------------------------------------ FAC_Comptes_Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Comptes_Clients")
    
    Call FAC_Comptes_Clients_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Comptes_Clients a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Comptes_Clients(r, readRows)
    
    'wshENC_Entête ------------------------------------------------------------ ENC_Entête
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "ENC_Entête")
    
    Call ENC_Entête_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "ENC_Entête a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_ENC_Entête(r, readRows)
    
    'wshENC_Détails ---------------------------------------------------------- ENC_Détails
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "ENC_Détails")
    
    Call ENC_Détails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "ENC_Détails a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_ENC_Détails(r, readRows)
    
    'wshFAC_Projets_Entête -------------------------------------------- FAC_Projets_Entête
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Projets_Entête")
    
    Call FAC_Projets_Entête_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Projets_Entête a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Projets_Entête(r, readRows)
    
    'wshFAC_Projets_Détails ------------------------------------------ FAC_Projets_Détails
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Projets_Détails")
    
    Call FAC_Projets_Détails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Projets_Détails a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Projets_Détails(r, readRows)
    
    'wshGL_Trans ---------------------------------------------------------------- GL_Trans
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "GL_Trans")
    
    Call GL_Trans_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "GL_Trans a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_GL_Trans(r, readRows)
    
    'wshTEC_TdB_Data -------------------------------------------------------- TEC_TdB_Data
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "TEC_TdB_Data")
    
    Call TEC_Import_All
    Call TEC_TdB_Update_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "TEC_TdB_Data a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_TEC_TdB_Data(r, readRows)
    
    'wshTEC_Local -------------------------------------------------------------- TEC_Local
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "TEC_Local")
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "TEC_Local a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    r = r + 1
    
    Call check_TEC(r, readRows)
    
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
    rng.value = "**** " & Format$(readRows, "###,##0") & _
                                    " lignes analysées dans l'ensemble des tables ***"
    rng.Characters(6, 6).Font.Color = vbRed
    rng.Characters(6, 6).Font.Bold = True
    r = r + 1
    
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "Vérification d'intégrité des tables"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    MsgBox "La vérification d'intégrité est terminé" & vbNewLine & vbNewLine & "Voir la feuille 'X_Analyse_Intégrité'", vbInformation
    
    ThisWorkbook.Worksheets("X_Analyse_Intégrité").Activate
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set rngToPrint = Nothing
    Set wsOutput = Nothing
    
    Call Log_Record("modAppli:Integrity_Verification", startTime)

End Sub

Private Sub check_Plan_Comptable(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_Plan_Comptable", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'dnrPlanComptable_All
    Dim arr As Variant
    Dim nbCol As Long
    nbCol = 4
    arr = Fn_Get_Plan_Comptable(nbCol) 'Returns array with 4 columns (Code, Description)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(UBound(arr, 1), "###,##0") & _
        " comptes et " & Format$(nbCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de 'dnr_PlanComptable_All'")
    r = r + 1
    
    If UBound(arr, 1) < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim dict_code_GL As New Dictionary
    Dim dict_descr_GL As New Dictionary
    
    Dim i As Long, codeGL As String, descrGL As String
    Dim GL_ID As Long
    Dim typeGL As String
    Dim cas_doublon_descr As Long, cas_doublon_code As Long, cas_type As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        codeGL = arr(i, 1)
        descrGL = arr(i, 2)
        If dict_descr_GL.Exists(descrGL) = False Then
            dict_descr_GL.Add descrGL, codeGL
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La description '" & descrGL & "' est un doublon pour le code de G/L '" & codeGL & "'")
            r = r + 1
            cas_doublon_descr = cas_doublon_descr + 1
        End If
        
        If dict_code_GL.Exists(codeGL) = False Then
            dict_code_GL.Add codeGL, descrGL
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le code de G/L '" & codeGL & "' est un doublon pour la description '" & descrGL & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
        
        GL_ID = arr(i, 3)
        typeGL = arr(i, 4)
        If InStr("Actifs^Passifs^Équité^Revenus^Dépenses^", typeGL) = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le type de compte '" & typeGL & "' est INVALIDE pour le code de G/L '" & codeGL & "'")
            r = r + 1
            cas_type = cas_type + 1
        End If
        
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " comptes ont été analysés!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_descr = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucun doublon de description")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_doublon_descr & " cas de doublons pour les descriptions")
        r = r + 1
    End If
    
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucun doublon de code de G/L")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_doublon_code & " cas de doublons pour les codes de G/L")
        r = r + 1
    End If
    
    If cas_type = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucun type de G/L invalide")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_type & " cas de types de G/L invalides")
        r = r + 1
    End If
    r = r + 1
    
Clean_Exit:
    'Libérer la mémoire
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_Plan_Comptable", startTime)

End Sub

Private Sub check_Clients(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_Clients", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'Fichier maître des Clients
    Dim ws As Worksheet: Set ws = wshBD_Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(ws.usedRange.rows.count - 1, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshBD_Clients'")
    r = r + 1
    
    Dim arr As Variant
    arr = wshBD_Clients.Range("A1").CurrentRegion.value
    If UBound(arr, 1) < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim dict_code_client As New Dictionary
    Dim dict_nom_client As New Dictionary
    
    Dim i As Long, code As String, nom As String, nomClientSysteme As String
    Dim eMail As String
    Dim cas_doublon_nom As Long
    Dim cas_doublon_code As Long
    Dim cas_doublon_nom_client_Systeme As Long
    Dim cas_courriel_invalide As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        nom = arr(i, 1)
        code = arr(i, 2)
        nomClientSysteme = arr(i, 3)
        
        'Doublon sur le nom ?
        If dict_nom_client.Exists(nom) = False Then
            dict_nom_client.Add nom, code
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "À la ligne " & i & ", le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        
        'Doublon sur le code de client ?
        If dict_code_client.Exists(code) = False Then
            dict_code_client.Add code, nom
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "À la ligne " & i & ", le code '" & code & "' est un doublon pour le client '" & nom & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
        
        If Trim(arr(i, 6)) <> "" Then
            If Fn_ValiderCourriel(arr(i, 6)) = False Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "À la ligne " & i & ", le courriel '" & arr(i, 6) & "' est INVALIDE pour le code '" & code & "'")
                r = r + 1
                cas_courriel_invalide = cas_courriel_invalide + 1
            End If
        End If
        
    Next i
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.value = "Un total de " & Format$(UBound(arr, 1) - 1, "##,##0") & " clients ont été analysés!"
    rng.Characters(13, 5).Font.Color = vbRed
    rng.Characters(13, 5).Font.Bold = True

    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucun doublon de nom")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        r = r + 1
    End If
    
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucun doublon de code")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        r = r + 1
    End If
    
    If cas_courriel_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Toutes les adresses courriel sont valides")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_courriel_invalide & " cas de courriels INVALIDES")
        r = r + 1
    End If
    
    r = r + 1
    
Clean_Exit:
    'Libérer la mémoire
    Set dict_code_client = Nothing
    Set dict_nom_client = Nothing
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_Clients", startTime)

End Sub

Private Sub check_Fournisseurs(ByRef r As Long, ByRef readRows As Long)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_Fournisseurs", 0)

    Application.ScreenUpdating = False

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshBD_fournisseurs
    Dim ws As Worksheet: Set ws = wshBD_Fournisseurs
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(ws.usedRange.rows.count - 1, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshBD_Fournisseurs'")
    r = r + 1
    
    Dim arr As Variant
    arr = wshBD_Fournisseurs.Range("A1").CurrentRegion.value
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        If dict_code_fournisseur.Exists(code) = False Then
            dict_code_fournisseur.Add code, nom
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le code '" & code & "' est un doublon pour le nom '" & nom & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
    Next i
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.value = "Un total de " & Format$(UBound(arr, 1) - 1, "#,##0") & " fournisseurs ont été analysés!"
    rng.Characters(13, 3).Font.Color = vbRed
    rng.Characters(13, 3).Font.Bold = True

    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucun doublon de nom")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        r = r + 1
    End If
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucun doublon de code")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        r = r + 1
    End If
    r = r + 1
    
Clean_Exit:
    'Libérer la mémoire
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_Fournisseurs", startTime)

End Sub

Private Sub check_ENC_Détails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_ENC_Détails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshENC_Détails
    Dim ws As Worksheet: Set ws = wshENC_Détails
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRowDetails As Long
    lastUsedRowDetails = ws.Cells(ws.rows.count, "A").End(xlUp).row
    If lastUsedRowDetails <= 2 - headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim col As Integer, nbCol As Integer
    col = 1
    'Boucle pour trouver la première colonne entièrement vide
    Do While col <= ws.columns.count
        If ws.Cells(1, col).value = "" Then
            nbCol = col
            Exit Do
        End If
        col = col + 1
    Loop
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRowDetails - headerRow, "###,##0") & _
        " lignes et " & Format$(nbCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    'ENC_Entête Worksheet
    Dim wsEntete As Worksheet: Set wsEntete = wshENC_Entête
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsEntete.Cells(wsEntete.rows.count, "A").End(xlUp).row
    Dim rngEntete As Range: Set rngEntete = wsEntete.Range("A2:A" & lastUsedRowEntete)
    Dim strPmtNo As String
    Dim i As Long
    For i = 2 To lastUsedRowEntete
        strPmtNo = strPmtNo & CLng(wsEntete.Range("A" & i).value) & "|"
    Next i
    
    'FAC_Entête Worksheet
    Dim wsFACEntete As Worksheet: Set wsFACEntete = wshFAC_Entête
    Dim lastUsedRowFacEntete As Long
    lastUsedRowFacEntete = wsFACEntete.Cells(wsFACEntete.rows.count, "A").End(xlUp).row
    Dim rngFACEntete As Range: Set rngFACEntete = wsFACEntete.Range("A2:A" & lastUsedRowFacEntete)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshENC_Détails'")
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
    For i = 2 To lastUsedRowDetails
        pmtNo = CLng(ws.Range("A" & i).value)
        If pmtNo <> oldpmtNo Then
            If InStr(strPmtNo, pmtNo) = 0 Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le paiement '" & pmtNo & "' à la ligne " & i & " n'existe pas dans ENC_Entête")
                r = r + 1
            End If
            strPmtNo = strPmtNo & pmtNo & "|"
            oldpmtNo = pmtNo
        End If
        
        Dim Inv_No As String
        Inv_No = CStr(ws.Range("B" & i).value)
        result = Application.WorksheetFunction.XLookup(Inv_No, _
                        rngFACEntete, _
                        rngFACEntete, _
                        "Not Found", _
                        0, _
                        1)
        If result = "Not Found" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "', ligne " & i & ", du paiement '" & pmtNo & "' n'existe pas dans FAC_Entête")
            r = r + 1
        End If
        
        If IsDate(ws.Range("D" & i).value) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La date '" & ws.Range("D" & i).value & "', ligne " & i & ", du paiment '" & pmtNo & "' est INVALIDE '")
            r = r + 1
        End If
        
        If IsNumeric(ws.Range("E" & i).value) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le montant '" & ws.Range("E" & i).value & "' du paiement '" & pmtNo & "' n'est pas numérique")
            r = r + 1
        Else
            If dictENC.Exists(Inv_No) Then
                dictENC(Inv_No) = dictENC(Inv_No) + ws.Range("E" & i).value
            Else
                dictENC.Add Inv_No, ws.Range("E" & i).value
            End If
            totalEncDetails = totalEncDetails + ws.Range("E" & i).value
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(lastUsedRowDetails - 1, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 1
    
    'Compare les encaissements accumulés (dictENC) avec wshFAC_Comptes_Clients
    Dim wsComptes_Clients As Worksheet: Set wsComptes_Clients = wshFAC_Comptes_Clients
    Dim lastUsedRow As Long
    lastUsedRow = wsComptes_Clients.Cells(wsComptes_Clients.rows.count, "A").End(xlUp).row
    Dim totalPaid As Currency
    
    For i = 3 To lastUsedRow
        Inv_No = wsComptes_Clients.Cells(i, 1).value
        totalPaid = wsComptes_Clients.Cells(i, "I").value
        If totalPaid <> dictENC(Inv_No) Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Pour la facture '" & Inv_No & "', le total des enc. " _
                            & "(wshFAC_Comptes_clients) " & Format$(totalPaid, "###,##0.00 $") _
                            & " est <> du détail des enc. " & Format$(dictENC(Inv_No), "###,##0.00 $"))
            r = r + 1
        End If
    Next i
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.value = "Total des encaissements : " & Format$(totalEncDetails, "##,###,##0.00 $")
    rng.Characters(InStr(rng.value, Left(totalEncDetails, 1)), 12).Font.Color = vbRed
    rng.Characters(InStr(rng.value, Left(totalEncDetails, 1)), 12).Font.Bold = True

    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + lastUsedRowDetails - 1
    
Clean_Exit:
    'Libérer la mémoire
    Set dictENC = Nothing
    Set rngEntete = Nothing
    Set rngFACEntete = Nothing
    Set ws = Nothing
    Set wsFACEntete = Nothing
    Set wsEntete = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_ENC_Détails", startTime)

End Sub

Private Sub check_ENC_Entête(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_ENC_Entête", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'Clients Master File
    Dim wsClients As Worksheet: Set wsClients = wshBD_Clients
    Dim lastUsedRowClient As Long
    lastUsedRowClient = wsClients.Cells(wsClients.rows.count, "B").End(xlUp).row
    Dim rngClients As Range: Set rngClients = wsClients.Range("B2:B" & lastUsedRowClient)
    
    'wshENC_Entête
    Dim ws As Worksheet: Set ws = wshENC_Entête
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim firstEmptyCol As Long
    firstEmptyCol = 1
    Do Until ws.Cells(headerRow, firstEmptyCol) = ""
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshENC_Entête'")
    r = r + 1
    
    If lastUsedRow = headerRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim arr As Variant
    arr = wshENC_Entête.Range("A1").CurrentRegion.Offset(1, 0) _
              .Resize(lastUsedRow - headerRow, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim pmtNo As String
    Dim totals As Currency
    Dim result As Variant
    For i = LBound(arr, 1) To UBound(arr, 1)
        pmtNo = arr(i, 1)
        If IsDate(arr(i, 2)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La date de paiement '" & arr(i, 2) & "' du paiement '" & arr(i, 1) & "' n'est pas VALIDE")
            r = r + 1
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le client '" & codeClient & "' du paiement '" & pmtNo & "' est INVALIDE")
            r = r + 1
        End If
        totals = totals + arr(i, 6)
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
    r = r + 1
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.value = "Total des encaissements : " & Format$(totals, "##,###,##0.00 $")
    rng.Characters(InStr(rng.value, Left(totals, 1)), 12).Font.Color = vbRed
    rng.Characters(InStr(rng.value, Left(totals, 1)), 12).Font.Bold = True
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
Clean_Exit:
    'Libérer la mémoire
    Set rngClients = Nothing
    Set ws = Nothing
    Set wsClients = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_ENC_Entête", startTime)

End Sub

Private Sub check_FAC_Détails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Détails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshFAC_Détails
    Dim ws As Worksheet: Set ws = wshFAC_Détails
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Entête
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsMaster.Cells(wsMaster.rows.count, "A").End(xlUp).row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A" & 1 + headerRow & ":A" & lastUsedRowEntete)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Détails'")
    r = r + 1
    
    'Transfer FAC_Details data from Worksheet into an Array (arr)
    Dim arr As Variant
    arr = wshFAC_Détails.Range("A1").CurrentRegion.Offset(1, 0).value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String, oldInv_No As String
    Dim result As Variant
    For i = LBound(arr, 1) + 2 To UBound(arr, 1) - 1 'Two lines of header !
        Inv_No = CStr(arr(i, 1))
'        Debug.Print "#887 - Inv_no = ", Inv_No, ", de type ", TypeName(Inv_No)
        If Inv_No <> oldInv_No Then
             result = Application.WorksheetFunction.XLookup(Inv_No, _
                                                    rngMaster, _
                                                    rngMaster, _
                                                    "Not Found", _
                                                    0, _
                                                    1)
            If result = "Not Found" Then
                Debug.Print "#895 - " & result
            End If
'            result = Application.WorksheetFunction.XLookup(ws.Cells(i, 1), rngMaster, rngMaster, "Not Found", 0, 1)
            oldInv_No = CStr(Inv_No)
        End If
        If result = "Not Found" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' à la ligne " & i & " n'existe pas dans FAC_Entête")
            r = r + 1
        End If
        If IsNumeric(arr(i, 3)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' à la ligne " & i & " le nombre d'heures est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' à la ligne " & i & " le taux horaire est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' à la ligne " & i & " le montant est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 2, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - 2
    
Clean_Exit:
    'Libérer la mémoire
    Set rngMaster = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Détails", startTime)

End Sub

Private Sub check_FAC_Entête(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Entête", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshFAC_Entête
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    Dim firstEmptyCol As Long
    firstEmptyCol = 1
    Do Until ws.Cells(headerRow, firstEmptyCol) = ""
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Entête'")
    r = r + 1
    
    If lastUsedRow = headerRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim arr As Variant
    arr = wshFAC_Entête.Range("A1").CurrentRegion.Offset(2, 0) _
              .Resize(lastUsedRow - headerRow, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String
    Dim totals(1 To 8, 1 To 2) As Currency
    Dim nbFactC As Long, nbFactAC As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Inv_No = arr(i, 1)
        If IsDate(arr(i, 2)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' à la ligne " & i & ", la date est de mauvais format '" & arr(i, 2) & "'")
                r = r + 1
            End If
        End If
        If arr(i, 3) <> "C" And arr(i, 3) <> "AC" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le type de facture '" & arr(i, 3) & "' pour la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If arr(i, 19) <> 0.09975 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le % de TVQ, soit '" & arr(i, 19) & "' pour la facture '" & Inv_No & "' est ERRONÉ")
            r = r + 1
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
    r = r + 1
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.value = "Totaux des factures CONFIRMÉES (" & nbFactC & " factures)"
    rng.Characters(InStr(rng.value, "CONFIRMÉES"), 10).Font.Color = vbRed
    rng.Characters(InStr(rng.value, "CONFIRMÉES"), 10).Font.Bold = True
    r = r + 1

    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.value = "       Total Fact. : " & Fn_Pad_A_String(Format$(totals(7, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.value, Left(totals(7, 1), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.value, Left(totals(7, 1), 1)), 15).Font.Bold = True
    r = r + 1

    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Acompte payé: " & _
            Fn_Pad_A_String(Format$(totals(8, 1), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 2
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.value = "Totaux des factures À CONFIRMER (" & nbFactAC & " factures)"
    rng.Characters(InStr(rng.value, "À CONFIRMER"), 11).Font.Color = vbRed
    rng.Characters(InStr(rng.value, "À CONFIRMER"), 11).Font.Bold = True
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.value = "       Total Fact. : " & Fn_Pad_A_String(Format$(totals(7, 2), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.value, Left(totals(7, 2), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.value, Left(totals(7, 2), 1)), 15).Font.Bold = True
    r = r + 1

    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Acompte payé: " & _
            Fn_Pad_A_String(Format$(totals(8, 2), "##,###,##0.00 $"), " ", 15, "L"))
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
Clean_Exit:
    'Libérer la mémoire
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Entête", startTime)

End Sub

Private Sub check_FAC_Comptes_Clients(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Comptes_Clients", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    Dim firstEmptyCol As Long
    firstEmptyCol = 1
    Do Until ws.Cells(headerRow, firstEmptyCol) = ""
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Comptes_Clients'")
    r = r + 1
    
    If lastUsedRow = headerRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    'Load every records into an Array
    Dim arr As Variant
    arr = wshFAC_Comptes_Clients.Range("A1").CurrentRegion.Offset(2, 0) _
              .Resize(lastUsedRow - headerRow, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String
    Dim totals(1 To 3, 1 To 2) As Currency
    Dim nbFactC As Long, nbFactAC As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Inv_No = arr(i, 1)
        Dim invType As String
        invType = Fn_Get_Invoice_Type(Inv_No)
        If invType <> "C" And invType <> "AC" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", le type de facture '" & invType & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsDate(CDate(arr(i, 2))) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", la date '" & arr(i, 2) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", la facture '" & Inv_No & "', la date est de mauvais format '" & arr(i, 2) & "'")
                r = r + 1
            End If
        End If
        If Fn_Validate_Client_Number(CStr(arr(i, 4))) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", le client '" & CStr(arr(i, 4)) & "' de la facture '" & Inv_No & "' est INVALIDE '")
            r = r + 1
        End If
        If arr(i, 5) <> "Paid" And arr(i, 5) <> "Unpaid" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", le statut '" & arr(i, 5) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsDate(CDate(arr(i, 7))) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", la date due '" & arr(i, 7) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", le total de la facture '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsNumeric(arr(i, 9)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", le montant payé à date '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsNumeric(arr(i, 10)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", le solde de la facture '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        'PLUG pour s'assurer que le solde impayé est belt et bien aligner sur le total et $ payé à date
        If arr(i, 10) <> arr(i, 8) - arr(i, 9) Then
            arr(i, 10) = arr(i, 8) - arr(i, 9)
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** À la ligne " & i + 2 & ", pour la facture '" & Inv_No & ", j'ai ajusté le solde de la facture à " & Format$(arr(i, 8), "###,##0.00 $") & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 11)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** L'âge (jours) de la facture '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If arr(i, 10) = 0 And arr(i, 5) = "Unpaid" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le statut '" & arr(i, 5) & "' de la facture '" & Inv_No & "', avec un solde de " & Format$(arr(i, 10), "#,##0.00 $") & " est INVALIDE")
            r = r + 1
        End If
        If arr(i, 10) <> 0 And arr(i, 5) = "Paid" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le statut '" & arr(i, 5) & "' de la facture '" & Inv_No & "', avec un solde de " & Format$(arr(i, 10), "#,##0.00 $") & " est INVALIDE")
            r = r + 1
        End If
        If invType = "C" Then
            totals(1, 1) = totals(1, 1) + arr(i, 8)
            totals(2, 1) = totals(2, 1) + arr(i, 9)
            totals(3, 1) = totals(3, 1) + arr(i, 10)
            nbFactC = nbFactC + 1
        Else
            totals(1, 2) = totals(1, 2) + arr(i, 8)
            totals(2, 2) = totals(2, 2) + arr(i, 9)
            totals(3, 2) = totals(3, 2) + arr(i, 10)
            nbFactAC = nbFactAC + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
    r = r + 1
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.value = "Totaux des factures CONFIRMÉES (" & nbFactC & " factures)"
    rng.Characters(InStr(rng.value, "CONFIRMÉES"), 10).Font.Color = vbRed
    rng.Characters(InStr(rng.value, "CONFIRMÉES"), 10).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.value = "       Total des factures        : " & Fn_Pad_A_String(Format$(totals(1, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.value, Left(totals(1, 1), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.value, Left(totals(1, 1), 1)), 15).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.value = "       Montants encaissés à date : " & Fn_Pad_A_String(Format$(totals(2, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.value, Left(totals(2, 1), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.value, Left(totals(2, 1), 1)), 15).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.value = "       Solde à recevoir          : " & Fn_Pad_A_String(Format$(totals(3, 1), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.value, Left(totals(3, 1), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.value, Left(totals(3, 1), 1)), 15).Font.Bold = True
    r = r + 2
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.value = "Total des factures À CONFIRMER (" & nbFactAC & " factures)"
    rng.Characters(InStr(rng.value, "À CONFIRMER"), 11).Font.Color = vbRed
    rng.Characters(InStr(rng.value, "À CONFIRMER"), 11).Font.Bold = True
    r = r + 1
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.value = "       Total des factures        : " & Fn_Pad_A_String(Format$(totals(1, 2), "##,###,##0.00 $"), " ", 15, "L")
    rng.Characters(InStr(rng.value, Left(totals(1, 2), 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.value, Left(totals(1, 2), 1)), 15).Font.Bold = True
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
Clean_Exit:
    'Libérer la mémoire
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Comptes_Clients", startTime)

End Sub

Private Sub check_FAC_Projets_Entête(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Projets_Entête", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Projets_Entête'")
    r = r + 1
    
    'Establish the number of rows before transferring it to an Array
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.rows.count
    If numRows <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Offset(1, 0).Resize(numRows - 1, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim projetID As String
    Dim codeClient As String
    For i = LBound(arr, 1) To UBound(arr, 1) 'One line of header !
        projetID = arr(i, 1)
        'Client valide ?
        codeClient = Trim(arr(i, 3))
        If Fn_Validate_Client_Number(codeClient) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " le Code de Client est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsDate(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 4) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " le total des honoraires est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 7)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du premier sommaire sont INVALIDES '" & arr(i, 7) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du premier sommaire est INVALIDE '" & arr(i, 8) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 9)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du premier sommaire sont INVALIDES '" & arr(i, 9) & "'")
            r = r + 1
        End If
        If arr(i, 11) <> "" And IsNumeric(arr(i, 11)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du second sommaire sont INVALIDES '" & arr(i, 11) & "'")
            r = r + 1
        End If
        If arr(i, 12) <> "" And IsNumeric(arr(i, 12)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du second sommaire est INVALIDE '" & arr(i, 12) & "'")
            r = r + 1
        End If
        If arr(i, 13) <> "" And IsNumeric(arr(i, 13)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du second sommaire sont INVALIDES '" & arr(i, 13) & "'")
            r = r + 1
        End If
        If arr(i, 15) <> "" And IsNumeric(arr(i, 15)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du troisième sommaire sont INVALIDES '" & arr(i, 15) & "'")
            r = r + 1
        End If
        If arr(i, 16) <> "" And IsNumeric(arr(i, 16)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du troisième sommaire est INVALIDE '" & arr(i, 16) & "'")
            r = r + 1
        End If
        If arr(i, 17) <> "" And IsNumeric(arr(i, 17)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du troisième sommaire sont INVALIDES '" & arr(i, 17) & "'")
            r = r + 1
        End If
        If arr(i, 19) <> "" And IsNumeric(arr(i, 19)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du quatrième sommaire sont INVALIDES '" & arr(i, 19) & "'")
            r = r + 1
        End If
        If arr(i, 20) <> "" And IsNumeric(arr(i, 20)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du quatrième sommaire est INVALIDE '" & arr(i, 20) & "'")
            r = r + 1
        End If
        If arr(i, 21) <> "" And IsNumeric(arr(i, 21)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du quatrième sommaire sont INVALIDES '" & arr(i, 21) & "'")
            r = r + 1
        End If
        If arr(i, 23) <> "" And IsNumeric(arr(i, 23)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les heures du cinquième sommaire sont INVALIDES '" & arr(i, 23) & "'")
            r = r + 1
        End If
        If arr(i, 24) <> "" And IsNumeric(arr(i, 24)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " le taux horaire du cinquième sommaire est INVALIDE '" & arr(i, 24) & "'")
            r = r + 1
        End If
        If arr(i, 25) <> "" And IsNumeric(arr(i, 25)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " les Honoraires du cinquième sommaire sont INVALIDES '" & arr(i, 25) & "'")
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " projets de factures a été analysés")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
Clean_Exit:
    'Libérer la mémoire
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Projets_Entête", startTime)

End Sub

Private Sub check_FAC_Projets_Détails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Projets_Détails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshFAC_Projets_Détails
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If

    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Projets_Entête
    lastUsedRow = wsMaster.Cells(wsMaster.rows.count, "A").End(xlUp).row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A2:A" & lastUsedRow)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshFAC_Projets_Détails'")
    r = r + 1
    
    'Transfer data from Worksheet into an Array (arr)
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.rows.count - 1 'Remove header
    If numRows < 1 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    'Charge le contenu de 'wshFAC_Projets_Détails' en mémoire (Array)
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Offset(1, 0).Resize(numRows, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim projetID As Long, oldProjetID As Long
    Dim codeClient As String
    Dim result As Variant
    
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le projet '" & projetID & "' à la ligne " & i & " n'existe pas dans FAC_Projets_Entête")
            r = r + 1
        End If
        'Client valide ?
        codeClient = Trim(arr(i, 3))
        If Fn_Validate_Client_Number(codeClient) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' à la ligne " & i & " le Code de Client est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le projet '" & projetID & "' à la ligne " & i & " le TECID est INVALIDE '" & arr(i, 4) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le projet '" & projetID & "' à la ligne " & i & " le ProfID est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le projet '" & projetID & "' à la ligne " & i & " les Heures sont INVALIDES '" & arr(i, 8) & "'")
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes ont été analysées")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
Clean_Exit:
    'Libérer la mémoire
    Set rngMaster = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Projets_Détails", startTime)

End Sub

Private Sub check_GL_Trans(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_GL_Trans", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshGL_Trans
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim firstEmptyCol As Long
    firstEmptyCol = 1
    Do Until ws.Cells(headerRow, firstEmptyCol) = ""
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshGL_Trans'")
    r = r + 1
    
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wshAdmin.Range("dnrPlanComptable_All")
    On Error GoTo 0

    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        r = r + 1
        Exit Sub
    End If
    
    Dim strCodeGL As String, strDescGL As String
    Dim ligne As Range
    For Each ligne In planComptable.rows
        strCodeGL = strCodeGL & ligne.Cells(1, 2).value & "|:|"
        strDescGL = strDescGL & ligne.Cells(1, 1).value & "|:|"
    Next ligne
    
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.rows.count - 1 'Remove the header row
    If numRows < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Offset(1, 0).Resize(numRows, ws.Range("A1").CurrentRegion.columns.count).value
    
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
    For i = LBound(arr, 1) To UBound(arr, 1)
        GL_Entry_No = arr(i, 1)
        If dict_GL_Entry.Exists(GL_Entry_No) = False Then
            dict_GL_Entry.Add GL_Entry_No, row
            sum_arr(row, 1) = GL_Entry_No
            row = row + 1
        End If
        If IsDate(arr(i, 2)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** L'écriture #  " & GL_Entry_No & " ' à la ligne " & i & " a une date INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** L'écriture #  " & GL_Entry_No & " ' à la ligne " & i & " a une date avec le mauvais format '" & arr(i, 2) & "'")
                r = r + 1
            End If
        End If
        glCode = arr(i, 5)
        If InStr(1, strCodeGL, glCode + "|:|") = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le compte '" & glCode & "' à la ligne " & i & " est INVALIDE '")
            r = r + 1
        End If
        If glCode = "1100" Then
            arTotal = arTotal + arr(i, 7) - arr(i, 8)
        End If
        glDescr = arr(i, 6)
        If InStr(1, strDescGL, glDescr + "|:|") = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La description du compte '" & glDescr & "' à la ligne " & i & " est INVALIDE")
            r = r + 1
        End If
        dt = arr(i, 7)
        If IsNumeric(dt) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le montant du débit '" & dt & "' à la ligne " & i & " n'est pas une valeur numérique")
            r = r + 1
        End If
        ct = arr(i, 8)
        If IsNumeric(ct) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le montant du débit '" & ct & "' à la ligne " & i & " n'est pas une valeur numérique")
            r = r + 1
        End If
        currentRow = dict_GL_Entry(GL_Entry_No)
        sum_arr(currentRow, 2) = sum_arr(currentRow, 2) + dt
        sum_arr(currentRow, 3) = sum_arr(currentRow, 3) + ct
        If arr(i, 10) <> "" Then
            If IsDate(arr(i, 10)) = False Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le TimeStamp '" & arr(i, 10) & "' à la ligne " & i & " n'est pas une date VALIDE")
                r = r + 1
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Écriture # " & v & " ne balance pas... Dt = " & Format$(dt, "###,###,##0.00") & " et Ct = " & Format$(ct, "###,###,##0.00"))
            r = r + 1
            cas_hors_balance = cas_hors_balance + 1
        End If
        sum_dt = sum_dt + dt
        sum_ct = sum_ct + ct
    Next v
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " lignes de transactions ont été analysées")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Un total de " & dict_GL_Entry.count & " écritures ont été analysées")
    r = r + 1
    
    If cas_hors_balance = 0 Then
        'Un peu de couleur
        Dim rng As Range: Set rng = wsOutput.Range("B" & r)
        rng.value = "       Chacune des écritures balancent au niveau de l'écriture"
        rng.Characters(InStr(rng.value, "C"), Len(rng.value) - 7).Font.Color = vbRed
        rng.Characters(InStr(rng.value, "C"), Len(rng.value) - 7).Font.Bold = True
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_hors_balance & " écriture(s) qui ne balance(nt) pas !!!")
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Les totaux des transactions sont:")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Dt = " & Format$(sum_dt, "###,###,##0.00 $"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Ct = " & Format$(sum_ct, "###,###,##0.00 $"))
    r = r + 1
    
    If sum_dt - sum_ct <> 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Hors-Balance de " & Format$(sum_dt - sum_ct, "###,###,##0.00 $"))
        r = r + 1
    End If
    
    'Un peu de couleur
    Set rng = wsOutput.Range("B" & r)
    rng.value = "Au Grand Livre, le solde des Comptes-Clients est de : " & Format$(arTotal, "##,###,##0.00 $")
    rng.Characters(InStr(rng.value, Left(arTotal, 1)), 15).Font.Color = vbRed
    rng.Characters(InStr(rng.value, Left(arTotal, 1)), 15).Font.Bold = True
    r = r + 2
    
Clean_Exit:
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set ligne = Nothing
    Set planComptable = Nothing
    Set v = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Call Log_Record("modAppli:check_GL_Trans", startTime)

End Sub

Private Sub check_TEC_TdB_Data(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_TEC_TdB_Data", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshTEC_TdB_Data
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim lastUsedCol As Long
    lastUsedCol = ws.Range("A2").End(xlToRight).Column
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshTEC_TdB_Data'")
    r = r + 1
    
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Offset(1)
    Dim dict_TEC_ID As New Dictionary
    Dim dict_prof As New Dictionary
    
    Dim i As Long, TECID As Long, profID As String, prof As String, dateTEC As Date, clientCode As String
    Dim minDate As Date, maxDate As Date
    Dim hres As Double, hres_non_detruites As Double
    Dim estDetruit As Boolean, estFacturable As Boolean, estFacturee As Boolean
    Dim cas_doublon_TECID As Long, cas_date_invalide As Long, cas_doublon_prof As Long, cas_doublon_client As Long
    Dim cas_hres_invalide As Long, cas_estFacturable_invalide As Long, cas_estFacturee_invalide As Long
    Dim cas_estDetruit_invalide As Long
    Dim total_hres_inscrites As Double, total_hres_detruites As Double, total_hres_facturees As Double
    Dim total_hres_facturable As Double, total_hres_TEC As Double, total_hres_non_facturable As Double
    
    minDate = "12/31/2999"
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        TECID = arr(i, 1)
        prof = arr(i, 3)
        dateTEC = arr(i, 4)
        If IsDate(dateTEC) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** TEC_ID =" & TECID & " a une date INVALIDE '" & dateTEC & " !!!")
            r = r + 1
            cas_date_invalide = cas_date_invalide + 1
        Else
            If dateTEC < minDate Then minDate = dateTEC
            If dateTEC > maxDate Then maxDate = dateTEC
        End If
        If dateTEC <> Int(dateTEC) Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** La date du TEC '" & dateTEC & "' n'est pas du bon format (H:M:S) pour le TEC_ID =" & TECID)
            r = r + 1
        End If
        clientCode = arr(i, 5)
        hres = arr(i, 8)
        If IsNumeric(hres) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            r = r + 1
            cas_hres_invalide = cas_hres_invalide + 1
        End If
        estFacturable = arr(i, 9)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            r = r + 1
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If
        estFacturee = arr(i, 10)
        If InStr("Vrai^Faux^", estFacturee & "^") = 0 Or Len(estFacturee) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            r = r + 1
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        estDetruit = arr(i, 11)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            r = r + 1
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
        If dict_TEC_ID.Exists(TECID) = False Then
            dict_TEC_ID.Add TECID, 0
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le TEC_ID '" & TECID & "' est un doublon pour la ligne '" & i & "'")
            r = r + 1
            cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        
        If dict_prof.Exists(prof & "-" & profID) = False Then
            dict_prof.Add prof & "-" & profID, 0
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " charges de temps ont été analysées!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
    If cas_doublon_TECID = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucun doublon de TEC_ID")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_doublon_TECID & " cas de doublons pour les TEC_ID")
        r = r + 1
    End If
    
    If cas_date_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune date INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    r = r + 1
    
    If cas_hres_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune heures INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturable_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune valeur 'estFacturable' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturee_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune valeur 'estFacturee' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        r = r + 1
    End If
    
    If cas_estDetruit_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune valeur 'estDetruit' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La somme des heures saisies donne ces résultats:")
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures SAISIES         : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures détruites       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures NETTES          : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "              Non_facturables : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "              Facturables     : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures facturées       : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable - total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.value = "       Heures TEC             : " & formattedHours
    rng.Characters(InStr(rng.value, ":") + 2, Len(formattedHours)).Font.Color = vbRed
    rng.Characters(InStr(rng.value, ":") + 2, Len(formattedHours)).Font.Bold = True
    r = r + 2

Clean_Exit:
    'Libérer la mémoire
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_TEC_TdB_Data", startTime)

End Sub

Private Sub check_TEC(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_TEC", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
'    Dim wsSommaire As Worksheet: Set wsSommaire = ThisWorkbook.Worksheets("X_Heures_Jour_Prof")
    
    Dim lastTECIDReported As Long
    lastTECIDReported = 2504 'What is the last TECID analyzed ?

    'wshTEC_Local
    Dim ws As Worksheet: Set ws = wshTEC_Local
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim lastUsedCol As Long
    lastUsedCol = ws.Range("A2").End(xlToRight).Column
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.Name & "' ou 'wshTEC_Local'")
    r = r + 1
    
    Dim rngCR As Range
    Set rngCR = ws.Range("A1").CurrentRegion
    Dim lastRow As Long, lastcol As Long
    lastRow = rngCR.rows.count
    lastcol = rngCR.columns.count
    Dim arr As Variant
    If lastRow > 2 Then
        'Décaler de 2 lignes (pour exclure les en-têtes) et redimensionner la plage
        arr = rngCR.Offset(2, 0).Resize(lastRow - 2, lastcol).value
'        arr = ws.Range("A1").CurrentRegion.Offset(2)
    Else
        MsgBox "Il n'y a aucune ligne de détail", vbInformation
        Exit Sub
    End If
    Dim dict_TEC_ID As New Dictionary
    Dim dict_prof As New Dictionary
    Dim dictFacture As New Dictionary
    Dim i As Long
    
    'Obtenir toutes les factures émises (wshFAC_Entête) et utiliser un dictionary pour les mémoriser
    Dim lastUsedRowFAC As Long
    lastUsedRowFAC = wshFAC_Entête.Cells(wshFAC_Entête.rows.count, "A").End(xlUp).row
    If lastUsedRowFAC > 2 Then
        For i = 3 To lastUsedRowFAC
            dictFacture.Add CStr(wshFAC_Entête.Cells(i, 1).value), 0
        Next i
    End If
    
    Dim TECID As Long, profID As String, prof As String, dateTEC As Date, dateFact As Date, testDate As Boolean
    Dim minDate As Date, maxDate As Date
    Dim maxTECID As Long
    Dim d As Integer, m As Integer, Y As Integer, p As Integer
    Dim codeClient As String, nomClient As String
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
    
'    Dim bigStrDateProf As String
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
    For i = LBound(arr, 1) To UBound(arr, 1)
        TECID = arr(i, 1)
        If TECID > maxTECID Then
            maxTECID = TECID
        End If
        'ProfessionnelID
        profID = arr(i, 2)
        'Professionnel
        prof = arr(i, 3)
        'Date
        dateTEC = arr(i, 4)
        testDate = IsDate(dateTEC)
        If testDate = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** TEC_ID =" & TECID & " a une date INVALIDE '" & dateTEC & " !!!")
            r = r + 1
            cas_date_invalide = cas_date_invalide + 1
        Else
            If dateTEC < minDate Then minDate = dateTEC
            If dateTEC > maxDate Then maxDate = dateTEC
        End If
        If dateTEC > Now() Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** TEC_ID =" & TECID & " a une date FUTURE '" & dateTEC & " !!!")
            r = r + 1
            cas_date_future = cas_date_future + 1
        End If
        If dateTEC <> Int(dateTEC) Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** La date du TEC '" & dateTEC & "' n'est pas du bon format (H:M:S) pour le TEC_ID =" & TECID)
            r = r + 1
        End If
        
        'Validate clientCode
        codeClient = Trim(arr(i, 5))
        If Fn_Validate_Client_Number(codeClient) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le code de client '" & codeClient & "' est INVALIDE !!!")
            r = r + 1
        End If
        nomClient = arr(i, 6)
        hres = arr(i, 8)
        testHres = IsNumeric(hres)
        If testHres = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            r = r + 1
            cas_hres_invalide = cas_hres_invalide + 1
        End If
        estFacturable = arr(i, 10)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            r = r + 1
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If

        'Analyse de la date de charge et du TimeStamp pour les dernières entrées
        If arr(i, 1) > lastTECIDReported Then
            'Date de la charge
            yy = year(arr(i, 4))
            mm = month(arr(i, 4))
            dd = day(arr(i, 4))
            strDict = Format$(DateSerial(yy, mm, dd), "yyyy-mm-dd") & " - " & _
                                Fn_Pad_A_String(CStr(arr(i, 3)), " ", 5, "R")
            If dictDateCharge.Exists(strDict) Then
                dictDateCharge(strDict) = dictDateCharge(strDict) + arr(i, 8)
            Else
                dictDateCharge.Add strDict, arr(i, 8)
            End If
            'TimeStamp
            yy = year(arr(i, 11))
            mm = month(arr(i, 11))
            dd = day(arr(i, 11))
            strDict = Format$(DateSerial(yy, mm, dd), "yyyy-mm-dd") & " - " & _
                                Fn_Pad_A_String(CStr(arr(i, 3)), " ", 5, "R")
            If dictTimeStamp.Exists(strDict) Then
                dictTimeStamp(strDict) = dictTimeStamp(strDict) + 1
            Else
                dictTimeStamp.Add strDict, 1
            End If
        End If

        estFacturee = UCase(arr(i, 12))
        If InStr("Vrai^VRAI^Faux^FAUX^", estFacturee & "^") = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            r = r + 1
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        
        If arr(i, 13) <> "" Then
            dateFact = arr(i, 13)
            testDate = IsDate(dateFact)
            If testDate = False Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** TEC_ID =" & TECID & " a une date de facture INVALIDE '" & dateFact & " !!!")
                r = r + 1
                cas_date_fact_invalide = cas_date_fact_invalide + 1
            End If
            If dateFact > Now() Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** TEC_ID =" & TECID & " a une date de facture FUTURE '" & dateFact & " !!!")
                r = r + 1
                cas_date_facture_future = cas_date_facture_future + 1
            End If
            If dateFact <> Int(dateFact) Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** La date de la facture '" & dateFact & "' n'est pas du bon format (H:M:S) pour le TEC_ID =" & TECID)
                r = r + 1
            End If
        End If
        
        estDetruit = arr(i, 14)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            r = r + 1
            cas_estDetruit_invalide = cas_estDetruit_invalide + 1
        End If
        
        invNo = CStr(arr(i, 16))
        If Len(invNo) > 0 Then
            If estFacturee <> "VRAI" Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & _
                    " - Incongruité entre le numéro de facture '" & invNo & "' et " & _
                    "'estFacture' qui vaut '" & estFacturee & "'")
                r = r + 1
            End If
            If dictFacture.Exists(invNo) = False Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & _
                    " - Le numéro de facture '" & invNo & "' " & _
                    "n'existe pas dans le fichier FAC_Entête")
                r = r + 1
            Else 'Accumule les heures pour cette facture
                dictFacture(invNo) = dictFacture(invNo) + arr(i, 8)
            End If
        Else
            If estFacturee = "Vrai" Or estFacturee = "VRAI" Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & _
                    " - Incongruité entre le numéro de facture vide et " & _
                    "'estFacture' qui vaut '" & estFacturee & "'")
                r = r + 1
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
            Debug.Print i & " Écart - " & TECID & " " & prof & " " & dateTEC & " " & h(1) & " " & h(2) & " vs. " & h(3) & " " & h(4)
            Stop
        End If
        
        'Dictionaries
        If dict_TEC_ID.Exists(TECID) = False Then
            dict_TEC_ID.Add TECID, 0
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le TEC_ID '" & TECID & "' est un doublon pour la ligne '" & i & "'")
            r = r + 1
            cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        If dict_prof.Exists(prof & "-" & profID) = False Then
            dict_prof.Add prof & "-" & profID, 0
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " charges de temps ont été analysées!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
    If cas_doublon_TECID = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucun doublon de TEC_ID")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_doublon_TECID & " cas de doublons pour les TEC_ID")
        r = r + 1
    End If
    
    If cas_date_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune date INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        r = r + 1
    End If
    
    If cas_date_future = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune date dans le futur")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_date_future & " cas de date FUTURE")
        r = r + 1
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    r = r + 1
    
    If cas_hres_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune heures INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturable_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune valeur 'estFacturable' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturee_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune valeur 'estFacturee' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        r = r + 1
    End If
    
    If cas_date_fact_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune date de facture INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_date_fact_invalide & " cas de date de facture INVALIDE")
        r = r + 1
    End If
    
    If cas_estDetruit_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune valeur 'estDetruit' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        r = r + 1
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Vérification des Heures Facturées par Facture")
    r = r + 1
    
    'Vérification des heures facturées selon 2 sources (TEC_Local vs. FAC_Détails)
    Dim key As Variant
    Dim totalHoursBilled As Double
    Dim cas_Heures_Differentes As Integer
    
    For Each key In dictFacture.keys
        totalHoursBilled = Fn_Get_TEC_Total_Invoice_AF(CStr(key), "Heures")
        If Round(totalHoursBilled, 2) <> Round(dictFacture(key), 2) Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Facture '" & CStr(key) & _
                    "', il y a un écart d'heures facturées entre TEC_Local & FAC_Détails - " & _
                        Round(dictFacture(key), 2) & " vs. " & Round(totalHoursBilled, 2))
            r = r + 1
            cas_Heures_Differentes = cas_Heures_Differentes + 1
        End If
    Next key

    If cas_Heures_Differentes = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Toutes les heures facturées balancent, selon les 2 sources")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Certaines factures sont à vérifier pour que les heures facturées balancent, selon les 2 sources")
        r = r + 1
    End If
        
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La somme des heures SAISIES donne ces résultats:")
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures SAISIES         : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures détruites       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures NETTES          : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "              Non_facturables : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "              Facturables     : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures facturées       : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable - total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    
    'Un peu de couleur
    Dim rng As Range: Set rng = wsOutput.Range("B" & r)
    rng.value = "       Heures TEC             : " & formattedHours
    rng.Characters(InStr(rng.value, ":") + 2, Len(formattedHours)).Font.Color = vbRed
    rng.Characters(InStr(rng.value, ":") + 2, Len(formattedHours)).Font.Bold = True
    r = r + 1
    
    Dim keys() As Variant
    
    'Tri & impression de dictDateCharge
    If dictDateCharge.count > 0 Then
        'Un peu de couleur
        Set rng = wsOutput.Range("B" & r)
        rng.value = "Sommaire des heures selon la DATE de la charge (" & maxTECID & ")"
        rng.Characters(InStr(rng.value, "(") + 1, Len(maxTECID)).Font.Color = vbGreen
        rng.Characters(InStr(rng.value, "(") + 1, Len(maxTECID)).Font.Bold = True
        r = r + 1
    
        keys = dictDateCharge.keys
        Call Fn_Quick_Sort(keys, LBound(keys), UBound(keys))
        'Parcourir les clés triées et afficher les heures
        For i = LBound(keys) To UBound(keys)
            key = keys(i)
            formattedHours = Format$(dictDateCharge(key), "#0.00")
            formattedHours = String(6 - Len(formattedHours), " ") & formattedHours
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       " & key & ":" & formattedHours & " heures")
            r = r + 1
        Next i
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Aucune nouvelle saisie d'heures (TECID > " & lastTECIDReported & ") ")
        r = r + 1
    End If
    
    'Tri & impression de dictTimeStamp
    If dictTimeStamp.count > 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Sommaire des heures saisies selon le 'TIMESTAMP'")
        r = r + 1
        keys = dictTimeStamp.keys
        Call Fn_Quick_Sort(keys, LBound(keys), UBound(keys))
        'Parcourir les clés triées et afficher les valeurs
        For i = LBound(keys) To UBound(keys)
            key = keys(i)
            formattedHours = Format$(dictTimeStamp(key), "##0")
            formattedHours = String(6 - Len(formattedHours), " ") & formattedHours
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       " & key & ":" & formattedHours & " entrée(s)")
            r = r + 1
'            Debug.Print "Clé: " & key & " - Valeur: " & dictTimeStamp(key)
        Next i
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Aucune nouvelle saisie d'heures (TECID > " & lastTECIDReported & ") ")
        r = r + 1
    End If
    r = r + 1
    
Clean_Exit:

    'Libérer la mémoire
    Set dictDateCharge = Nothing
    Set dictTimeStamp = Nothing
    Set dict_TEC_ID = Nothing
    Set rngCR = Nothing
    Set ws = Nothing
'    Set wsSommaire = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_TEC", startTime)

End Sub

'CommentOut - 2024-11-14
'Sub ADMIN_DataFiles_Folder_Selection() '2024-03-28 @ 14:10
'
'    Dim SharedFolder As FileDialog: Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
'
'    With SharedFolder
'        .Title = "Choisir le répertoire de données partagées, selon les instructions de l'Administrateur"
'        .AllowMultiSelect = False
'        If .show = -1 Then
'            wshAdmin.Range("F5").value = .selectedItems(1)
'        End If
'    End With
'
'    'Libérer la mémoire
'    Set SharedFolder = Nothing
'
'End Sub

'CommentOut - 2024-11-14
'Sub ADMIN_Invoices_Excel_Folder_Selection() '2024-08-04 @ 07:30
'
'    Dim SharedFolder As FileDialog: Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
'
'    With SharedFolder
'        .Title = "Choisir le répertoire des factures (Format Excel)"
'        .AllowMultiSelect = False
'        If .show = -1 Then
'            wshAdmin.Range("F7").value = .selectedItems(1)
'        End If
'    End With
'
'    'Libérer la mémoire
'    Set SharedFolder = Nothing
'
'End Sub
'

Sub Make_It_As_Header(r As Range)

    With r
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
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
    ws.columns.AutoFit
    
    'Libérer la mémoire
    Set r = Nothing
    Set ws = Nothing

End Sub

Sub Add_Message_To_WorkSheet(ws As Worksheet, r As Long, c As Long, m As String)

    ws.Cells(r, c).value = m
    If c = 1 Then
        ws.Cells(r, c).Font.Bold = True
    End If

End Sub

'CommentOut - 2024-11-14
'Sub ADMIN_PDF_Folder_Selection() '2024-03-28 @ 14:10
'
'    Dim PDFFolder As FileDialog: Set PDFFolder = Application.FileDialog(msoFileDialogFolderPicker)
'
'    With PDFFolder
'        .Title = "Choisir le répertoire des copies de facture (PDF), selon les instructions de l'Administrateur"
'        .AllowMultiSelect = False
'        If .show = -1 Then
'            wshAdmin.Range("F6").value = .selectedItems(1)
'        End If
'    End With
'
'    'Libérer la mémoire
'    Set PDFFolder = Nothing
'
'End Sub
'
Sub Apply_Conditional_Formatting_Alternate(rng As Range, headerRows As Long, Optional EmptyLine As Boolean = False)

    'Avons-nous un Range valide ?
    If rng Is Nothing Or rng.rows.count <= headerRows Then
        Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = rng.Worksheet
    Dim dataRange As Range
    
   ' Définir la plage de données à laquelle appliquer la mise en forme conditionnelle, en
    'excluant les lignes d'en-tête
    Set dataRange = rng.Resize(rng.rows.count - headerRows).Offset(headerRows, 0)
    
    'Effacer les formats conditionnels existants sur la plage de données
    dataRange.Interior.ColorIndex = xlNone

    'Appliquer les couleurs en alternance
    Dim i As Long
    For i = 1 To dataRange.rows.count
        'Vérifier la position réelle de la ligne dans la feuille
        If (dataRange.rows(i).row + headerRows) Mod 2 = 0 Then
            dataRange.rows(i).Interior.Color = RGB(173, 216, 230) ' Bleu pâle
        End If
    Next i
    
    'Libérer la mémoire
    Set dataRange = Nothing
    Set ws = Nothing
    
End Sub

Sub Apply_Worksheet_Format(ws As Worksheet, rng As Range, headerRow As Long)

    'Common stuff to all worksheets
    rng.EntireColumn.AutoFit 'Autofit all columns
    
    'Conditional Formatting (many steps)
    '1) Remove existing conditional formatting
        rng.Cells.FormatConditions.Delete 'Remove the worksheet conditional formatting
    
    '2) Define the usedRange to data only (exclude header row(s))
        Dim numRows As Long
        numRows = rng.CurrentRegion.rows.count - headerRow
        Dim usedRange As Range
        If numRows > 0 Then
            On Error Resume Next
            Set usedRange = rng.Offset(headerRow, 0).Resize(numRows, rng.columns.count)
            On Error GoTo 0
        End If
    
    '3) Add the standard conditional formatting
        If Not usedRange Is Nothing Then
            With usedRange
                .FormatConditions.Add Type:=xlExpression, _
                    Formula1:="=MOD(LIGNE();2)=1"
                .FormatConditions(.FormatConditions.count).SetFirstPriority
                With .FormatConditions(1).Font
                    .Strikethrough = False
                    .TintAndShade = 0
                End With
                With .FormatConditions(1).Interior
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
                End With
                .FormatConditions(1).StopIfTrue = False
            End With
        Else
            MsgBox "usedRange is Nothing!"
        End If
        
    'Specific formats to worksheets
    Dim lastUsedRow As Long
    lastUsedRow = rng.rows.count
    If lastUsedRow = headerRow Then
        Exit Sub
    End If
    
    Dim firstDataRow As Long
    firstDataRow = headerRow + 1
    
    Select Case rng.Worksheet.CodeName
        Case "wshBD_Clients"
            
        Case "wshBD_Fournisseurs"
            
        Case "wshDEB_Recurrent"
            With wshDEB_Recurrent
                .Range("A2:M" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("C2:C" & lastUsedRow & _
                     ", D2:D" & lastUsedRow & _
                     ", E2:E" & lastUsedRow & _
                     ", G2:G" & lastUsedRow).HorizontalAlignment = xlLeft
                With .Range("I2:N" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
                End With
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
                .Range("A2:B" & lastUsedRow & ", " & _
                       "D2:G" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("C3:C" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("H3:J" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("B3:B" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("G3:G" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("H3:J" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("J3").formula = "=H3-I3"
                .Range("K3").formula = "=NOW()-G3"
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
                      .Range("C2:C" & lastUsedRow)).HorizontalAlignment = xlCenter
                Union(.Range("B2:B" & lastUsedRow), _
                      .Range("D2:D" & lastUsedRow), _
                      .Range("G2:G" & lastUsedRow)).HorizontalAlignment = xlLeft
                With .Range("E2:F" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
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
'CommentOut - 2024-11-14
'                With .Range("A2:A" & lastUsedRow) _
'                    .Range("J2:J" & lastUsedRow).Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
'                    .ThemeColor = xlThemeColorAccent5
'                    .TintAndShade = 0.799981688894314
'                    .PatternTintAndShade = 0
'                End With
            End With
        
        Case "wshTEC_Local"
            With wshTEC_Local
                .Range("A2:P" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("F2:F" & lastUsedRow & ", G2:G" & lastUsedRow & ", I2:I" & lastUsedRow & _
                            ", O2:O" & lastUsedRow).HorizontalAlignment = xlLeft
                            
                .Range("H2:H" & lastUsedRow).NumberFormat = "#0.00"
                .Range("D2:D" & lastUsedRow).NumberFormat = "yyyy-mm-dd"
                .Range("K2:K" & lastUsedRow).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                .columns("F").ColumnWidth = 40
                .columns("G").ColumnWidth = 55
                .columns("I").ColumnWidth = 20
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

Sub Get_TEC_Pour_Deplacements()  '2024-09-05 @ 10:22

    'Mise en place de la feuille de sortie (output)
    Dim strOutput As String
    strOutput = "X_TEC_Déplacements"
    Call CreateOrReplaceWorksheet(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").value = "Date"
    wsOutput.Range("B1").value = "Date"
    wsOutput.Range("C1").value = "Nom du client"
    wsOutput.Range("D1").value = "Heures"
    wsOutput.Range("E1").value = "Adresse_1"
    wsOutput.Range("F1").value = "Adresse_2"
    wsOutput.Range("G1").value = "Ville"
    wsOutput.Range("H1").value = "Province"
    wsOutput.Range("I1").value = "CodePostal"
    wsOutput.Range("J1").value = "DistanceKM"
    wsOutput.Range("K1").value = "Montant"
    Call Make_It_As_Header(wsOutput.Range("A1:K1"))
    
    'Feuille pour les clients
    Dim wsMF As Worksheet: Set wsMF = wshBD_Clients
    Dim lastUsedRowClientMF As Long
    lastUsedRowClientMF = wsMF.Cells(wsMF.rows.count, "A").End(xlUp).row
    Dim rngClientsMF As Range
    Set rngClientsMF = wsMF.Range("A1:A" & lastUsedRowClientMF)
    
    'Get From and To Dates
    Dim dateFrom As Date, dateTo As Date
    dateFrom = wshAdmin.Range("MoisPrecDe").value
    dateTo = wshAdmin.Range("MoisPrecA").value
    
    'Analyse de TEC_Local
    Call TEC_Import_All
    
    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    Dim lastUsedRowTEC As Long
    lastUsedRowTEC = wsTEC.Cells(wsTEC.rows.count, "A").End(xlUp).row
    
    Dim rowOutput As Long
    rowOutput = 2 'Skip the header
    Dim clientData As Variant
    Dim i As Long
    For i = 3 To lastUsedRowTEC
        If wsTEC.Cells(i, 3).value = "GC" And _
            wsTEC.Cells(i, 4).value >= dateFrom And _
            wsTEC.Cells(i, 4).value <= dateTo And _
            UCase(wsTEC.Cells(i, 14).value) <> "VRAI" Then
                wsOutput.Cells(rowOutput, 1).value = CDate(wsTEC.Cells(i, 4).value)
                wsOutput.Cells(rowOutput, 2).value = CDate(wsTEC.Cells(i, 4).value)
                wsOutput.Cells(rowOutput, 4).value = wsTEC.Cells(i, 8).value
                clientData = Fn_Rechercher_Client_Par_ID(Trim(wsTEC.Cells(i, 5).value), wsMF)
                If IsArray(clientData) Then
                    wsOutput.Cells(rowOutput, 3).value = clientData(1, fClntMFClientNom)
                    wsOutput.Cells(rowOutput, 5).value = clientData(1, fClntMFAdresse_1)
                    wsOutput.Cells(rowOutput, 6).value = clientData(1, fClntMFAdresse_2)
                    wsOutput.Cells(rowOutput, 7).value = clientData(1, fClntMFVille)
                    wsOutput.Cells(rowOutput, 8).value = clientData(1, fClntMFProvince)
                    wsOutput.Cells(rowOutput, 9).value = clientData(1, fClntMFCodePostal)
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
    
    wsOutput.columns.AutoFit

    'Améliore le Look (saute 1 ligne entre chaque jour)
    For i = rowOutput To 3 Step -1
        If Len(Trim(wsOutput.Cells(i, 3).value)) > 0 Then
            If wsOutput.Cells(i, 2).value <> wsOutput.Cells(i - 1, 2).value Then
                wsOutput.rows(i).Insert Shift:=xlDown
                wsOutput.Cells(i, 1).value = wsOutput.Cells(i - 1, 2).value
            End If
        End If
    Next i
    
    rowOutput = wsOutput.Cells(wsOutput.rows.count, "A").End(xlUp).row
    
    'Améliore le Look (cache la date, le client et l'adresse si deux charges & +)
    Dim base As String
    For i = 2 To rowOutput
        If i = 2 Then
            base = wsOutput.Cells(i, 2).value & wsOutput.Cells(i, 3).value
        End If
        If i > 2 And Len(wsOutput.Cells(i, 2).value) > 0 Then
            If wsOutput.Cells(i, 2).value & wsOutput.Cells(i, 3).value = base Then
                wsOutput.Cells(i, 2).value = ""
                wsOutput.Cells(i, 3).value = ""
                wsOutput.Cells(i, 5).value = ""
                wsOutput.Cells(i, 6).value = ""
                wsOutput.Cells(i, 7).value = ""
                wsOutput.Cells(i, 8).value = ""
                wsOutput.Cells(i, 9).value = ""
            Else
                base = wsOutput.Cells(i, 2).value & wsOutput.Cells(i, 3).value
            End If
        End If
    Next i
    
    'Result print setup - 2024-08-05 @ 05:16
    rowOutput = wsOutput.Cells(wsOutput.rows.count, "A").End(xlUp).row
    
    For i = 3 To rowOutput
        If wsOutput.Cells(i, 1).value > wsOutput.Cells(i - 1, 1).value Then
            wsOutput.Cells(i, 2).Font.Bold = True
        Else
            wsOutput.Cells(i, 2).value = ""
        End If
    Next i
    'Première date est en caractère gras
    wsOutput.Cells(2, 2).Font.Bold = True
    rowOutput = rowOutput + 2
    wsOutput.Range("A" & rowOutput).value = "**** " & Format$(lastUsedRowTEC - 2, "###,##0") & _
                                        " charges de temps analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("B2:K" & rowOutput)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

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

Sub Search_Unclean_Set()

    Dim ws As Worksheet: ' Set ws = Feuil4
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "B").End(xlUp).row
    
    Dim strSet As String
    Dim strForEach As String
    Dim strNothing As String
    Dim code As String
    Dim saveModule As String
    Dim saveLineNo As String
    Dim saveProcedure As String
    Dim wsOutput As Worksheet: ' Set wsOutput = Feuil3
    
    Dim i As Long
    Dim j As Long
    Dim r As Long
    
    For i = 2 To lastUsedRow
        If saveModule = "" Then
            saveModule = ws.Cells(i, 3)
            saveLineNo = ws.Cells(i + 1, 4)
            saveProcedure = ws.Cells(i + 1, 5)
        End If
        If i = 1232 Then Stop
        'On change de procédure
        If ws.Cells(i, 2) = "" Then
            If strSet <> "" Or strForEach <> "" Then
                If strSet <> "" Then
                    Dim arrSet() As String
                    arrSet = Split(strSet, "|")
                    For j = 0 To UBound(arrSet, 1) - 1
                        If InStr(strNothing, arrSet(j) & "|") = 0 Then
                            r = r + 1
                            wsOutput.Cells(r, 1) = i
                            wsOutput.Cells(r, 2) = saveModule
                            wsOutput.Cells(r, 3) = saveProcedure
                            wsOutput.Cells(r, 4) = saveLineNo
                            wsOutput.Cells(r, 5) = arrSet(j)
                            wsOutput.Cells(r, 6) = strNothing
                        End If
                    Next j
                End If
                If strForEach <> "" Then
                    Dim arrForEach() As String
                    arrForEach = Split(strForEach, "|")
                    For j = 0 To UBound(arrForEach, 1) - 1
                        If InStr(strNothing, arrForEach(j) & "|") = 0 Then
                            r = r + 1
                            wsOutput.Cells(r, 1) = i
                            wsOutput.Cells(r, 2) = saveModule
                            wsOutput.Cells(r, 3) = saveProcedure
                            wsOutput.Cells(r, 4) = saveLineNo
                            wsOutput.Cells(r, 5) = arrForEach(j)
                            wsOutput.Cells(r, 6) = strNothing
                        End If
                    Next j
                End If
            End If
            strSet = ""
            strForEach = ""
            strNothing = ""
            saveModule = ws.Cells(i + 1, 3)
            saveLineNo = ws.Cells(i + 1, 4)
            saveProcedure = ws.Cells(i + 1, 5)
        Else
            code = ws.Cells(i, 6)
            If InStr(code, "Set ") = 1 And InStr(code, " = Nothing") > 0 Then
                strNothing = strNothing & Mid(code, 5, Len(code) - 14) & "|"
            Else
                code = Replace(code, "RecordSet", "recordset")
                code = Replace(code, "Property Set", "Property set")
                If InStr(code, "Set ") > 0 Then
                    strSet = strSet & Mid(code, InStr(code, "Set ") + 4, InStr(Mid(code, InStr(code, "Set ")), " = ") - 5) & "|"
                Else
                    If InStr(code, "For Each") > 0 Then
                        strForEach = strForEach & Mid(code, InStr(code, "For Each ") + 9, InStr(Mid(code, InStr(code, "For Each ") + 9), " ") - 1) & "|"
                    End If
                End If
            End If
        End If
    Next i
    
    MsgBox "Traiement terminé " & i
    
End Sub

Sub Dynamic_Range_Redefine_Plan_Comptable() '2024-07-04 @ 10:39
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:Dynamic_Range_Redefine_Plan_Comptable", 0)

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



