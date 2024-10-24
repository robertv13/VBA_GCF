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

    'Cleaning memory - 2024-07-01 @ 09:34
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

    'Lock the checkbox
    rng.Locked = True
    
    'Protect the worksheet
    With rng.Parent
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

End Sub

Public Sub UnprotectCells(rng As Range)

    'Lock the checkbox
    rng.Locked = False
    
    'Protect the worksheet
    With rng.Parent
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

End Sub

Sub Update_Hres_Jour_Prof() '2024-08-15 @ 06:30

    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets("X_Heures_Jour_Prof")
    
    Dim wsTgt As Worksheet
    Set wsTgt = ThisWorkbook.Worksheets("TEC_Hres_Jour_Prof")
    
    Dim lastUsedRowSrc As Long
    lastUsedRowSrc = wsSrc.Cells(wsSrc.rows.count, "A").End(xlUp).Row '2024-08-15 @ 06:17
    
    wsTgt.Range("A2:I" & wsTgt.Cells(wsTgt.rows.count, "A").End(xlUp).Row).ClearContents
    
    'Copy columns A to H (from Source to Target), using Copy and Paste Special
    wsSrc.Range("A2:I" & lastUsedRowSrc).Copy
    wsTgt.Cells(2, 1).PasteSpecial Paste:=xlPasteValues
    
    'Clear the clipboard
    Application.CutCopyMode = False
    
    Call Update_Pivot_Table
    
    MsgBox "L'importation des Heures par Jour / Professionnel est compl�t�e" & _
            vbNewLine & vbNewLine & "Ainsi que la mise � jour du Pivot Table", _
            vbExclamation

    'Clean up
    Set wsSrc = Nothing
    Set wsTgt = Nothing
    
End Sub

Sub Update_Pivot_Table() '2024-08-15 @ 06:34

    'Define the worksheet containing the data
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("TEC_Hres_Jour_Prof")
    
    'Find the last row of your data
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    
    'Define the new data range
    Dim rngData As Range: Set rngData = ws.Range("A1:I" & lastUsedRow)
    
    'Update the Pivot Table
    Dim pt As pivotTable: Set pt = ws.PivotTables("ptHresJourProf")
    pt.ChangePivotCache ThisWorkbook.PivotCaches.Create( _
                        SourceType:=xlDatabase, _
                        SourceData:=rngData)
    
    'Refresh the Pivot Table
    pt.RefreshTable
    
    'Let go the Objects
    Set pt = Nothing
    Set rngData = Nothing
    Set ws = Nothing

End Sub

Public Sub ArrayToRange(ByRef Data As Variant _
                        , ByVal outRange As Range _
                        , Optional ByVal clearExistingData As Boolean = True _
                        , Optional ByVal clearExistingHeaderSize As Long = 1)
                        
    If clearExistingData = True Then
        outRange.CurrentRegion.Offset(clearExistingHeaderSize).ClearContents
    End If
    
    Dim rows As Long, columns As Long
    rows = UBound(Data, 1) - LBound(Data, 1) + 1
    columns = UBound(Data, 2) - LBound(Data, 2) + 1
    outRange.Resize(rows, columns).value = Data
    
End Sub

Sub CreateOrReplaceWorksheet(wsName As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli_Utils:CreateOrReplaceWorksheet", 0)
    
    'Check if the worksheet exists
    Dim ws As Worksheet
    Dim wsExists As Boolean
    For Each ws In ThisWorkbook.Worksheets
        wsExists = False
        If ws.name = wsName Then
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
    ws.name = wsName

    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
    Call Log_Record("modAppli_Utils:CreateOrReplaceWorksheet", startTime)

End Sub

Sub Detect_Circular_References_In_Workbook() '2024-07-24 @ 07:31
    
    Dim circRef As String
    circRef = ""
    Dim circRefCount As Long
    circRefCount = 0
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Dim formulaCells As Range
        Set formulaCells = ws.usedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        
        Dim cell As Range
        Dim cellCount As Long
        If Not formulaCells Is Nothing Then
            For Each cell In formulaCells
                On Error Resume Next
                cellCount = Application.CircularReference.count
                On Error GoTo 0
                
                If cellCount > 0 Then
                    circRef = circRef & ws.name & "!" & cell.Address & vbCrLf
                    circRefCount = circRefCount + 1
                End If
            Next cell
        End If
    Next ws
    
    If circRefCount > 0 Then
        MsgBox "Il existe des r�f�rences circulaires dans le Workbook dans les cellules suivantes:" & vbCrLf & circRef, vbExclamation
    Else
        MsgBox "Il n'existe aucune r�f�rence circulaire dans ce Workbook .", vbInformation
    End If
    
    'Clean up
    Set cell = Nothing
    Set formulaCells = Nothing
    Set ws = Nothing
    
End Sub

Public Sub Integrity_Verification() '2024-07-06 @ 12:56

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:Integrity_Verification", 0)

    Application.ScreenUpdating = False
    
    Call Erase_And_Create_Worksheet("X_Analyse_Int�grit�")
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    wsOutput.Range("A1").value = "Feuille"
    wsOutput.Range("B1").value = "Message"
    wsOutput.Range("C1").value = "TimeStamp"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"))

'    Call Erase_And_Create_Worksheet("X_Heures_Jour_Prof")
'    Dim wsSommaire As Worksheet: Set wsSommaire = ThisWorkbook.Worksheets("X_Heures_Jour_Prof")
'    wsSommaire.Range("A1").value = "Date"
'    wsSommaire.Range("B1").value = "Prof."
'    wsSommaire.Range("C1").value = "H/Saisies"
'    wsSommaire.Range("D1").value = "H/D�truites"
'    wsSommaire.Range("E1").value = "H/Nettes"
'    wsSommaire.Range("F1").value = "H/NFact"
'    wsSommaire.Range("G1").value = "H/Fact"
'    wsSommaire.Range("H1").value = "H/Factur�es"
'    wsSommaire.Range("I1").value = "H/TEC"
'    Call Make_It_As_Header(wsSommaire.Range("A1:I1"))

    'Data starts at row 2
    Dim r As Long: r = 2
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "R�pertoire utilis�")
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, wshAdmin.Range("FolderSharedData").value & DATA_PATH)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1

    'Fichier utilis�
    Dim masterFileName As String
    masterFileName = "GCF_BD_MASTER.xlsx"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Fichier utilis�")
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, masterFileName)
    r = r + 1
    
    'Date derni�re modification du fichier Ma�tre
    Dim fullFileName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & DATA_PATH & Application.PathSeparator & masterFileName
    Dim ddm As Date
    Dim j As Long, h As Long, m As Long, s As Long
    Call Get_Date_Derniere_Modification(fullFileName, ddm, j, h, m, s)
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Date dern. modification")
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, Format$(ddm, "dd-mm-yyyy hh:mm:ss") & _
            " soit " & j & " jours, " & h & " heures, " & m & " minutes et " & s & " secondes")
    r = r + 2
    
    Dim readRows As Long
    
    'dnrPlanComptable ----------------------------------------------------- Plan Comptable
    Application.ScreenUpdating = True
    Application.EnableEvents = False
    wshMenu.Range("H29").value = "V�rification du Plan Comptable"
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Plan Comptable")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    
    Call check_Plan_Comptable(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshBD_Clients --------------------------------------------------------------- Clients
    wshMenu.Range("H29").value = "V�rification des donn�es des clients"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "BD_Clients")
    
    Call Client_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_Clients(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshBD_Fournisseurs ----------------------------------------------------- Fournisseurs
    wshMenu.Range("H29").value = "V�rification des donn�es des fournisseurs"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "BD_Fournisseurs")
    
    Call Fournisseur_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_Fournisseurs(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshENC_D�tails ---------------------------------------------------------- ENC_D�tails
    wshMenu.Range("H29").value = "V�rification du d�tail des encaissements"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "ENC_D�tails")
    
    Call ENC_D�tails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "ENC_D�tails a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_ENC_D�tails(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshENC_Ent�te ------------------------------------------------------------ ENC_Ent�te
    wshMenu.Range("H29").value = "V�rification des ent�tes des encaissements"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "ENC_Ent�te")
    
    Call ENC_Ent�te_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "ENC_Ent�te a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_ENC_Ent�te(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshFAC_D�tails ---------------------------------------------------------- FAC_D�tails
    wshMenu.Range("H29").value = "V�rification des d�tails des factures"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_D�tails")
    
    Call FAC_D�tails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_D�tails a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_D�tails(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshFAC_Ent�te ------------------------------------------------------------ FAC_Ent�te
    wshMenu.Range("H29").value = "V�rification des ent�tes des factures"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Ent�te")
    
    Call FAC_Ent�te_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Ent�te a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Ent�te(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshFAC_Comptes_Clients ------------------------------------------ FAC_Comptes_Clients
    wshMenu.Range("H29").value = "V�rification des comptes clients"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Comptes_Clients")
    
    Call FAC_Comptes_Clients_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Comptes_Clients a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Comptes_Clients(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshFAC_Projets_D�tails ------------------------------------------ FAC_Projets_D�tails
    wshMenu.Range("H29").value = "V�rification des ent�tes de projets de factures"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Projets_D�tails")
    
    Call FAC_Projets_D�tails_Import_All
    Call FAC_Projets_Ent�te_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Projets_D�tails a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Projets_D�tails(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshFAC_Projets_Ent�te -------------------------------------------- FAC_Projets_Ent�te
    wshMenu.Range("H29").value = "V�rification des d�tails de projets de factures"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Projets_Ent�te")
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Projets_Ent�te a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Projets_Ent�te(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshGL_Trans ---------------------------------------------------------------- GL_Trans
    wshMenu.Range("H29").value = "V�rification des transactions du Grand Livre"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "GL_Trans")
    
    Call GL_Trans_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "GL_Trans a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))

    Call check_GL_Trans(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshTEC_TdB_Data -------------------------------------------------------- TEC_TdB_Data
    wshMenu.Range("H29").value = "V�rification des donn�es de tableau de bord (TEC)"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "TEC_TdB_Data")
    
    Call TEC_Import_All
    Call TEC_TdB_Update_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "TEC_Local a �t� import�e du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    
    Call check_TEC_TdB_Data(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'wshTEC_Local -------------------------------------------------------------- TEC_Local
    wshMenu.Range("H29").value = "V�rification des TEC"
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "TEC_Local")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_TEC(r, readRows)
    wshMenu.Range("H29").value = ""
    
    'Adjust the Output Worksheet
    With wsOutput.Range("A2:C" & r).Font
        .name = "Courier New"
        .size = 10
    End With
    
    wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
   'Result print setup - 2024-07-20 @ 14:31
    Dim lastUsedRow As Long
    lastUsedRow = r
    wsOutput.Range("A" & lastUsedRow).value = "**** " & Format$(readRows, "###,##0") & _
                                    " lignes analys�es dans l'ensemble des tables ***"
    
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "V�rification d'int�grit� des tables"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    MsgBox "La v�rification d'int�grit� est termin�" & vbNewLine & vbNewLine & "Voir la feuille 'X_Analyse_Int�grit�'", vbInformation
    
    ThisWorkbook.Worksheets("X_Analyse_Int�grit�").Activate
    
    Application.ScreenUpdating = True
    
    'Clean up
    Set rngToPrint = Nothing
    Set wsOutput = Nothing
    
    Call Log_Record("modAppli:Integrity_Verification", startTime)

End Sub

Private Sub check_Clients(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_Clients", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshBD_Clients
    Dim ws As Worksheet: Set ws = wshBD_Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(ws.usedRange.rows.count - 1, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshBD_Clients'")
    r = r + 1
    
    Dim arr As Variant
    arr = wshBD_Clients.Range("A1").CurrentRegion.value
    If UBound(arr, 1) < 2 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim dict_code_client As New Dictionary
    Dim dict_nom_client As New Dictionary
    
    Dim i As Long, code As String, nom As String
    Dim cas_doublon_nom As Long
    Dim cas_doublon_code As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        nom = arr(i, 1)
        code = arr(i, 2)
        
        If dict_nom_client.Exists(nom) = False Then
            dict_nom_client.Add nom, code
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "� la ligne " & i & ", le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        
        If dict_code_client.Exists(code) = False Then
            dict_code_client.Add code, nom
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "� la ligne " & i & ", le code '" & code & "' est un doublon pour le client '" & nom & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
        
    Next i
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 1, "##,##0") & " clients ont �t� analys�s!")
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
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_Clients", startTime)

End Sub

Private Sub check_Fournisseurs(ByRef r As Long, ByRef readRows As Long)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_Fournisseurs", 0)

    Application.ScreenUpdating = False

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshBD_fournisseurs
    Dim ws As Worksheet: Set ws = wshBD_Fournisseurs
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(ws.usedRange.rows.count - 1, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshBD_Fournisseurs'")
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
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 1, "#,##0") & " fournisseurs ont �t� analys�s!")
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
    'Cleaning memory - 2024-07-04 @ 12:37
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_Fournisseurs", startTime)

End Sub

Private Sub check_ENC_D�tails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_ENC_D�tails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshENC_D�tails
    Dim ws As Worksheet: Set ws = wshENC_D�tails
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRowDetails As Long
    lastUsedRowDetails = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastUsedRowDetails <= 2 - headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim col As Integer, nbCol As Integer
    col = 1
    'Boucle pour trouver la premi�re colonne enti�rement vide
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
    
    'ENC_Ent�te Worksheet
    Dim wsEntete As Worksheet: Set wsEntete = wshENC_Ent�te
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsEntete.Cells(wsEntete.rows.count, "A").End(xlUp).Row
    Dim rngEntete As Range: Set rngEntete = wsEntete.Range("A2:A" & lastUsedRowEntete)
    Dim strPmtNo As String
    Dim i As Long
    For i = 2 To lastUsedRowEntete
        strPmtNo = strPmtNo & wsEntete.Range("A" & i).value & "|"
    Next i
    
    'FAC_Ent�te Worksheet
    Dim wsFACEntete As Worksheet: Set wsFACEntete = wshFAC_Ent�te
    Dim lastUsedRowFacEntete As Long
    lastUsedRowFacEntete = wsFACEntete.Cells(wsFACEntete.rows.count, "A").End(xlUp).Row
    Dim rngFACEntete As Range: Set rngFACEntete = wsFACEntete.Range("A2:A" & lastUsedRowFacEntete)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshENC_D�tails'")
    r = r + 1
    
    'Array pointer
    Dim Row As Long: Row = 1
    Dim currentRow As Long
        
    Dim pmtNo As Long, oldpmtNo As Long
    Dim result As Variant
    Dim totalEncDetails As Currency
    For i = 2 To lastUsedRowDetails
        pmtNo = ws.Range("A" & i).value
        If pmtNo <> oldpmtNo Then
            If InStr(strPmtNo, pmtNo) = 0 Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le paiement '" & pmtNo & "' � la ligne " & i & " n'existe pas dans ENC_Ent�te")
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' du paiement '" & pmtNo & "' n'existe pas dans FAC_Ent�te")
            r = r + 1
        End If
        
        If IsDate(ws.Range("D" & i).value) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La date '" & ws.Range("D" & i).value & "' du paiment '" & pmtNo & "' est INVALIDE '")
            r = r + 1
        End If
        
        If IsNumeric(ws.Range("E" & i).value) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le montant '" & ws.Range("E" & i).value & "' du paiement '" & pmtNo & "' n'est pas num�rique")
            r = r + 1
        End If
        totalEncDetails = totalEncDetails + ws.Range("E" & i).value
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Total des encaissements : " & Format$(totalEncDetails, "#,##0.00 $"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(lastUsedRowDetails - 1, "##,##0") & " lignes de transactions ont �t� analys�es")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + lastUsedRowDetails - 1
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rngEntete = Nothing
    Set rngFACEntete = Nothing
    Set ws = Nothing
    Set wsFACEntete = Nothing
    Set wsEntete = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_ENC_D�tails", startTime)

End Sub

Private Sub check_ENC_Ent�te(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_ENC_Ent�te", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'Clients Master File
    Dim wsClients As Worksheet: Set wsClients = wshBD_Clients
    Dim lastUsedRowClient As Long
    lastUsedRowClient = wsClients.Cells(wsClients.rows.count, "B").End(xlUp).Row
    Dim rngClients As Range: Set rngClients = wsClients.Range("B2:B" & lastUsedRowClient)
    
    'wshENC_Ent�te
    Dim ws As Worksheet: Set ws = wshENC_Ent�te
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).Row
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshENC_Ent�te'")
    r = r + 1
    
    If lastUsedRow = headerRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim arr As Variant
    arr = wshENC_Ent�te.Range("A1").CurrentRegion.Offset(1, 0) _
              .Resize(lastUsedRow - headerRow, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim Row As Long: Row = 1
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Total des encaissements : " & Format$(totals, "#,##0.00 $"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont �t� analys�es")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rngClients = Nothing
    Set ws = Nothing
    Set wsClients = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_ENC_Ent�te", startTime)

End Sub

Private Sub check_FAC_D�tails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_D�tails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshFAC_D�tails
    Dim ws As Worksheet: Set ws = wshFAC_D�tails
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Ent�te
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsMaster.Cells(wsMaster.rows.count, "A").End(xlUp).Row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A" & 1 + headerRow & ":A" & lastUsedRowEntete)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_D�tails'")
    r = r + 1
    
    'Transfer FAC_Details data from Worksheet into an Array (arr)
    Dim arr As Variant
    arr = wshFAC_D�tails.Range("A1").CurrentRegion.Offset(1, 0).value
    
    'Array pointer
    Dim Row As Long: Row = 1
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' � la ligne " & i & " n'existe pas dans FAC_Ent�te")
            r = r + 1
        End If
        If IsNumeric(arr(i, 3)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' � la ligne " & i & " le nombre d'heures est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' � la ligne " & i & " le taux horaire est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' � la ligne " & i & " le montant est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 2, "##,##0") & " lignes de transactions ont �t� analys�es")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - 2
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rngMaster = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_D�tails", startTime)

End Sub

Private Sub check_FAC_Ent�te(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Ent�te", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshFAC_Ent�te
    Dim ws As Worksheet: Set ws = wshFAC_Ent�te
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).Row
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Ent�te'")
    r = r + 1
    
    If lastUsedRow = headerRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim arr As Variant
    arr = wshFAC_Ent�te.Range("A1").CurrentRegion.Offset(2, 0) _
              .Resize(lastUsedRow - headerRow, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim Row As Long: Row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String
    Dim totals(1 To 8, 1 To 2) As Currency
    Dim nbFactC As Long, nbFactAC As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Inv_No = arr(i, 1)
        If IsDate(arr(i, 2)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' � la ligne " & i & " la date est INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' � la ligne " & i & ", la date est de mauvais format '" & arr(i, 2) & "'")
                r = r + 1
            End If
        End If
        If arr(i, 3) <> "C" And arr(i, 3) <> "AC" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le type de facture '" & arr(i, 3) & "' pour la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If arr(i, 19) <> 0.09975 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le % de TVQ, soit '" & arr(i, 19) & "' pour la facture '" & Inv_No & "' est ERRON�")
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont �t� analys�es")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Factures CONFIRM�ES (" & nbFactC & " factures)")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Total Fact. : " & _
            Fn_Pad_A_String(Format$(totals(7, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Acompte pay�: " & _
            Fn_Pad_A_String(Format$(totals(8, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 2
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Factures � CONFIRMER (" & nbFactAC & " factures)")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 2), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 2), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 2), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 2), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 2), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 2), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Total Fact. : " & _
            Fn_Pad_A_String(Format$(totals(7, 2), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Acompte pay�: " & _
            Fn_Pad_A_String(Format$(totals(8, 2), "##,##0.00 $"), " ", 14, "L"))
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Ent�te", startTime)

End Sub

Private Sub check_FAC_Comptes_Clients(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Comptes_Clients", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).Row
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Comptes_Clients'")
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
    Dim Row As Long: Row = 1
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le type de facture '" & invType & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsDate(CDate(arr(i, 2))) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La date '" & arr(i, 2) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La facture '" & Inv_No & "' � la ligne " & i & ", la date est de mauvais format '" & arr(i, 2) & "'")
                r = r + 1
            End If
        End If
        If Fn_Validate_Client_Number(CStr(arr(i, 4))) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le client '" & CStr(arr(i, 4)) & "' de la facture '" & Inv_No & "' est INVALIDE '")
            r = r + 1
        End If
        If arr(i, 5) <> "Paid" And arr(i, 5) <> "Unpaid" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le statut '" & arr(i, 5) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsDate(CDate(arr(i, 7))) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La date due '" & arr(i, 7) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le total de la facture '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsNumeric(arr(i, 9)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le montant pay� � date '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsNumeric(arr(i, 10)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le solde de la facture '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        'PLUG pour s'assurer que le solde impay� est belt et bien aligner sur le total et $ pay� � date
        If arr(i, 10) <> arr(i, 8) - arr(i, 9) Then
            arr(i, 10) = arr(i, 8) - arr(i, 9)
        End If
        If IsNumeric(arr(i, 11)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** L'�ge (jours) de la facture '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont �t� analys�es")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Factures CONFIRM�ES (" & nbFactC & " factures)")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Total des factures        : " & Fn_Pad_A_String(Format$(totals(1, 1), "###,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Montants encaiss�s � date : " & Fn_Pad_A_String(Format$(totals(2, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Solde � recevoir          : " & Fn_Pad_A_String(Format$(totals(3, 1), "##,##0.00 $"), " ", 14, "L"))
    r = r + 2
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Factures � CONFIRMER (" & nbFactAC & " factures)")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Total des factures        : " & Fn_Pad_A_String(Format$(totals(1, 2), "###,##0.00 $"), " ", 14, "L"))
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Comptes_Clients", startTime)

End Sub

Private Sub check_FAC_Projets_D�tails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Projets_D�tails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshFAC_Projets_D�tails
    Dim ws As Worksheet: Set ws = wshFAC_Projets_D�tails
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If

    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Projets_Ent�te
    lastUsedRow = wsMaster.Cells(wsMaster.rows.count, "A").End(xlUp).Row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A2:A" & lastUsedRow)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Projets_D�tails'")
    r = r + 1
    
    'Transfer data from Worksheet into an Array (arr)
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.rows.count - 1 'Remove header
    If numRows < 1 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Offset(1, 0).Resize(numRows, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim Row As Long: Row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim projetID As Long, oldProjetID As Long
    Dim codeClient As String
    Dim lookUpValue As Long, result As Variant
    For i = LBound(arr, 1) To UBound(arr, 1)
        projetID = CLng(arr(i, 1))
        lookUpValue = projetID
        If projetID <> oldProjetID Then
            If projetID = 4 Then Stop
            result = Application.WorksheetFunction.XLookup(lookUpValue, _
                                                           rngMaster, _
                                                           rngMaster, _
                                                           "Not Found", _
                                                           0, _
                                                           1)
            oldProjetID = projetID
        End If
        If result = "Not Found" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le projet '" & projetID & "' � la ligne " & i & " n'existe pas dans FAC_Projets_Ent�te")
            r = r + 1
        End If
        'Client valide ?
        codeClient = Trim(arr(i, 3))
        If Fn_Validate_Client_Number(codeClient) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " le Code de Client est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le projet '" & projetID & "' � la ligne " & i & " le TECID est INVALIDE '" & arr(i, 4) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le projet '" & projetID & "' � la ligne " & i & " le ProfID est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le projet '" & projetID & "' � la ligne " & i & " les Heures sont INVALIDES '" & arr(i, 8) & "'")
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes ont �t� analys�es")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rngMaster = Nothing
    Set ws = Nothing
    Set wsMaster = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Projets_D�tails", startTime)

End Sub

Private Sub check_FAC_Projets_Ent�te(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Projets_Ent�te", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Ent�te
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Projets_Ent�te'")
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
    Dim Row As Long: Row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim projetID As String
    Dim codeClient As String
    For i = LBound(arr, 1) To UBound(arr, 1) 'One line of header !
        projetID = arr(i, 1)
        'Client valide ?
        codeClient = Trim(arr(i, 3))
        If Fn_Validate_Client_Number(codeClient) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " le Code de Client est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsDate(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " la date est INVALIDE '" & arr(i, 4) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " le total des honoraires est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 7)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les heures du premier sommaire sont INVALIDES '" & arr(i, 7) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " le taux horaire du premier sommaire est INVALIDE '" & arr(i, 8) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 9)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les Honoraires du premier sommaire sont INVALIDES '" & arr(i, 9) & "'")
            r = r + 1
        End If
        If arr(i, 11) <> "" And IsNumeric(arr(i, 11)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les heures du second sommaire sont INVALIDES '" & arr(i, 11) & "'")
            r = r + 1
        End If
        If arr(i, 12) <> "" And IsNumeric(arr(i, 12)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " le taux horaire du second sommaire est INVALIDE '" & arr(i, 12) & "'")
            r = r + 1
        End If
        If arr(i, 13) <> "" And IsNumeric(arr(i, 13)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les Honoraires du second sommaire sont INVALIDES '" & arr(i, 13) & "'")
            r = r + 1
        End If
        If arr(i, 15) <> "" And IsNumeric(arr(i, 15)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les heures du troisi�me sommaire sont INVALIDES '" & arr(i, 15) & "'")
            r = r + 1
        End If
        If arr(i, 16) <> "" And IsNumeric(arr(i, 16)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " le taux horaire du troisi�me sommaire est INVALIDE '" & arr(i, 16) & "'")
            r = r + 1
        End If
        If arr(i, 17) <> "" And IsNumeric(arr(i, 17)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les Honoraires du troisi�me sommaire sont INVALIDES '" & arr(i, 17) & "'")
            r = r + 1
        End If
        If arr(i, 19) <> "" And IsNumeric(arr(i, 19)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les heures du quatri�me sommaire sont INVALIDES '" & arr(i, 19) & "'")
            r = r + 1
        End If
        If arr(i, 20) <> "" And IsNumeric(arr(i, 20)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " le taux horaire du quatri�me sommaire est INVALIDE '" & arr(i, 20) & "'")
            r = r + 1
        End If
        If arr(i, 21) <> "" And IsNumeric(arr(i, 21)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les Honoraires du quatri�me sommaire sont INVALIDES '" & arr(i, 21) & "'")
            r = r + 1
        End If
        If arr(i, 23) <> "" And IsNumeric(arr(i, 23)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les heures du cinqui�me sommaire sont INVALIDES '" & arr(i, 23) & "'")
            r = r + 1
        End If
        If arr(i, 24) <> "" And IsNumeric(arr(i, 24)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " le taux horaire du cinqui�me sommaire est INVALIDE '" & arr(i, 24) & "'")
            r = r + 1
        End If
        If arr(i, 25) <> "" And IsNumeric(arr(i, 25)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Dans le projet '" & projetID & "' � la ligne " & i & " les Honoraires du cinqui�me sommaire sont INVALIDES '" & arr(i, 25) & "'")
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " projets de factures a �t� analys�s")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Projets_Ent�te", startTime)

End Sub

Private Sub check_GL_Trans(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_GL_Trans", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshGL_Trans
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshGL_Trans'")
    r = r + 1
    
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wshAdmin.Range("dnrPlanComptable_All")
    On Error GoTo 0

    If planComptable Is Nothing Then
        MsgBox "La plage nomm�e 'dnrPlanComptable_All' n'a pas �t� trouv�e ou est INVALIDE!", vbExclamation
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La plage nomm�e 'dnrPlanComptable_All' n'a pas �t� trouv�e!")
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
    Dim Row As Long: Row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim dt As Currency, ct As Currency
    Dim arTotal As Currency
    Dim GL_Entry_No As String, glCode As String, glDescr As String
    Dim result As Variant
    For i = LBound(arr, 1) To UBound(arr, 1)
        GL_Entry_No = arr(i, 1)
        If dict_GL_Entry.Exists(GL_Entry_No) = False Then
            dict_GL_Entry.Add GL_Entry_No, Row
            sum_arr(Row, 1) = GL_Entry_No
            Row = Row + 1
        End If
        If IsDate(arr(i, 2)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** L'�criture #  " & GL_Entry_No & " ' � la ligne " & i & " a une date INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
        Else
            If arr(i, 2) <> Int(arr(i, 2)) Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** L'�criture #  " & GL_Entry_No & " ' � la ligne " & i & " a une date avec le mauvais format '" & arr(i, 2) & "'")
                r = r + 1
            End If
        End If
        glCode = arr(i, 5)
        If InStr(1, strCodeGL, glCode + "|:|") = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le compte '" & glCode & "' � la ligne " & i & " est INVALIDE '")
            r = r + 1
        End If
        If glCode = "1100" Then
            arTotal = arTotal + arr(i, 7) - arr(i, 8)
        End If
        glDescr = arr(i, 6)
        If InStr(1, strDescGL, glDescr + "|:|") = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** La description du compte '" & glDescr & "' � la ligne " & i & " est INVALIDE")
            r = r + 1
        End If
        dt = arr(i, 7)
        If IsNumeric(dt) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le montant du d�bit '" & dt & "' � la ligne " & i & " n'est pas une valeur num�rique")
            r = r + 1
        End If
        ct = arr(i, 8)
        If IsNumeric(ct) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le montant du d�bit '" & ct & "' � la ligne " & i & " n'est pas une valeur num�rique")
            r = r + 1
        End If
        currentRow = dict_GL_Entry(GL_Entry_No)
        sum_arr(currentRow, 2) = sum_arr(currentRow, 2) + dt
        sum_arr(currentRow, 3) = sum_arr(currentRow, 3) + ct
        If arr(i, 10) <> "" Then
            If IsDate(arr(i, 10)) = False Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Le TimeStamp '" & arr(i, 10) & "' � la ligne " & i & " n'est pas une date VALIDE")
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** �criture # " & v & " ne balance pas... Dt = " & Format$(dt, "###,###,##0.00") & " et Ct = " & Format$(ct, "###,###,##0.00"))
            r = r + 1
            cas_hors_balance = cas_hors_balance + 1
        End If
        sum_dt = sum_dt + dt
        sum_ct = sum_ct + ct
    Next v
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " lignes de transactions ont �t� analys�es")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Un total de " & dict_GL_Entry.count & " �critures ont �t� analys�es")
    r = r + 1
    
    If cas_hors_balance = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Chacune des �critures balancent au niveau de l'�criture")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_hors_balance & " �criture(s) qui ne balance(nt) pas !!!")
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Au Grand Livre, le solde des Comptes-Clients est de : " & Format$(arTotal, "###,###,##0.00 $"))
    r = r + 2
    
Clean_Exit:
    Application.ScreenUpdating = True
    
    'Clean up
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
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
    'wshTEC_TdB_Data
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshTEC_TdB_Data'")
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
        
        'Heures d�truites
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
        
        'Heures FACTUR�ES
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " charges de temps ont �t� analys�es!")
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
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La somme des heures saisies donne ces r�sultats:")
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
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures d�truites       : " & formattedHours)
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
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures factur�es       : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable - total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures TEC             : " & formattedHours)
    r = r + 2

Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_TEC_TdB_Data", startTime)

End Sub

Private Sub check_Plan_Comptable(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_Plan_Comptable", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
    
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
        If InStr("Actifs^Passifs^�quit�^Revenus^D�penses^", typeGL) = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le type de compte '" & typeGL & "' est INVALIDE pour le code de G/L '" & codeGL & "'")
            r = r + 1
            cas_type = cas_type + 1
        End If
        
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " comptes ont �t� analys�s!")
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
    'Cleaning memory - 2024-07-01 @ 09:34
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_Plan_Comptable", startTime)

End Sub

Private Sub check_TEC(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_TEC", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Int�grit�")
'    Dim wsSommaire As Worksheet: Set wsSommaire = ThisWorkbook.Worksheets("X_Heures_Jour_Prof")
    
    Dim lastTECIDReported As Long
    lastTECIDReported = 1917 'What is the last TECID analyzed ?

    'wshTEC_Local
    Dim ws As Worksheet: Set ws = wshTEC_Local
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshTEC_Local'")
    r = r + 1
    
    Dim rngCR As Range
    Set rngCR = ws.Range("A1").CurrentRegion
    Dim lastRow As Long, lastcol As Long
    lastRow = rngCR.rows.count
    lastcol = rngCR.columns.count
    Dim arr As Variant
    If lastRow > 2 Then
        'D�caler de 2 lignes (pour exclure les en-t�tes) et redimensionner la plage
        arr = rngCR.Offset(2, 0).Resize(lastRow - 2, lastcol).value
'        arr = ws.Range("A1").CurrentRegion.Offset(2)
    Else
        MsgBox "Il n'y a aucune ligne de d�tail", vbInformation
        Exit Sub
    End If
    Dim dict_TEC_ID As New Dictionary
    Dim dict_prof As New Dictionary
    Dim dictFacture As New Dictionary
    Dim i As Long
    
    'Obtenir toutes les factures �mises et utiliser un dictionary pour les m�moriser
    Dim lastUsedRowFAC As Long
    lastUsedRowFAC = wshFAC_Ent�te.Cells(wshFAC_Ent�te.rows.count, "A").End(xlUp).Row
    If lastUsedRowFAC > 2 Then
        For i = 3 To lastUsedRowFAC
            dictFacture.Add CStr(wshFAC_Ent�te.Cells(i, 1).value), 0
        Next i
    End If
    
    Dim TECID As Long, profID As String, prof As String, dateTEC As Date, testDate As Boolean
    Dim minDate As Date, maxDate As Date
    Dim maxTECID As Long
    Dim D As Integer, m As Integer, Y As Integer, p As Integer
    Dim codeClient As String, nomClient As String
    Dim isClientValid As Boolean
    Dim hres As Double, testHres As Boolean, estFacturable As Boolean
    Dim estFacturee As Boolean, estDetruit As Boolean
    Dim invNo As String
    Dim cas_doublon_TECID As Long, cas_date_invalide As Long, cas_doublon_prof As Long, cas_doublon_client As Long
    Dim cas_date_future As Long
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
        profID = arr(i, 2)
        prof = arr(i, 3)
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

        'Analyse de la date de charge et du TimeStamp pour les derni�res entr�es
        If arr(i, 1) > lastTECIDReported Then
'            Debug.Print "#2135: "; i; " "; arr(i, 1); " "; arr(i, 4); " ", arr(i, 6); " "; arr(i, 8)
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
                    " - Incongruit� entre le num�ro de facture '" & invNo & "' et " & _
                    "'estFacture' qui vaut '" & estFacturee & "'")
                r = r + 1
            End If
            If dictFacture.Exists(invNo) = False Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & _
                    " - Le num�ro de facture '" & invNo & "' " & _
                    "n'existe pas dans le fichier FAC_Ent�te")
                r = r + 1
            End If
        Else
            If estFacturee = "Vrai" Or estFacturee = "VRAI" Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** TEC_ID = " & TECID & _
                    " - Incongruit� entre le num�ro de facture vide et " & _
                    "'estFacture' qui vaut '" & estFacturee & "'")
                r = r + 1
            End If
        End If

        'Accumule les heures
        Dim h(1 To 6) As Double
        
        'Heures INSCRITES
        total_hres_inscrites = total_hres_inscrites + hres
        h(1) = hres
        
        'Heures D�TRUITES
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
        
        'Heures FACTUR�ES
        h(5) = 0
        If estFacturee = "Vrai" And Fn_Is_Client_Facturable(codeClient) = True Then
                total_hres_facturees = total_hres_facturees + hres
                h(5) = hres
        End If
        
        'Heures TEC = Heures Facturables - Heures factur�es
        If h(3) Then
            h(6) = h(3) - h(5)
        Else
            h(6) = 0
        End If
        
        If h(1) - h(2) <> h(3) + h(4) Then
            Debug.Print i & " �cart - " & TECID & " " & prof & " " & dateTEC & " " & h(1) & " " & h(2) & " vs. " & h(3) & " " & h(4)
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
        
        'Summary by Date
'        D = day(dateTEC)
'        m = month(dateTEC)
'        Y = year(dateTEC)
'        keyDate = Format$(Y, "0000") & Format$(m, "00") & Format$(D, "00") & Fn_Pad_A_String(prof, " ", 4, "L")
'        p = InStr(bigStrDateProf, keyDate)
'        If p = 0 Then
'            rArr = rArr + 1
'            pArr = rArr
'            bigStrDateProf = bigStrDateProf & keyDate & Format$(rArr, "0000") & "|"
'        Else
'            pArr = Mid(bigStrDateProf, p + 12, 4)
'        End If
'        arrHres(pArr, 1) = arrHres(pArr, 1) + h(1)
'        arrHres(pArr, 2) = arrHres(pArr, 2) + h(2)
'        arrHres(pArr, 3) = arrHres(pArr, 3) + h(3)
'        arrHres(pArr, 4) = arrHres(pArr, 4) + h(4)
'        arrHres(pArr, 5) = arrHres(pArr, 5) + h(5)
'        arrHres(pArr, 6) = arrHres(pArr, 6) + h(6)
    Next i
    
'    Call SortDelimitedString(bigStrDateProf, "|")
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " charges de temps ont �t� analys�es!")
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
    
    If cas_estDetruit_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Aucune valeur 'estDetruit' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "****** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        r = r + 1
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La somme des heures SAISIES donne ces r�sultats:")
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures SAISIES        :  " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures d�truites       : " & formattedHours)
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
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures factur�es       : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable - total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       Heures TEC             : " & formattedHours)
    r = r + 1

    Dim keys() As Variant
    Dim key As Variant
    
    'Tri & impression de dictDateCharge
    If dictDateCharge.count > 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Sommaire des heures par DATE de la charge (" & maxTECID & ")")
        r = r + 1
        keys = dictDateCharge.keys
        Call Fn_Quick_Sort(keys, LBound(keys), UBound(keys))
        'Parcourir les cl�s tri�es et afficher les heures
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
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Sommaire des heures saisies par 'TIMESTAMP' (" & maxTECID & ")")
        r = r + 1
        keys = dictTimeStamp.keys
        Call Fn_Quick_Sort(keys, LBound(keys), UBound(keys))
        'Parcourir les cl�s tri�es et afficher les valeurs
        For i = LBound(keys) To UBound(keys)
            key = keys(i)
            formattedHours = Format$(dictTimeStamp(key), "##0")
            formattedHours = String(6 - Len(formattedHours), " ") & formattedHours
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "       " & key & ":" & formattedHours & " entr�e(s)")
            r = r + 1
'            Debug.Print "Cl�: " & key & " - Valeur: " & dictTimeStamp(key)
        Next i
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Aucune nouvelle saisie d'heures (TECID > " & lastTECIDReported & ") ")
        r = r + 1
    End If
    r = r + 1
    
'    Dim r2 As Integer
'    r2 = 2 'Output to wsSommaire
    
'    Dim components() As String
'    components = Split(bigStrDateProf, "|")
'
'    Dim dateStr As String
'    For i = LBound(components) To UBound(components)
'        dateStr = Left(components(i), 8)
'        dateStr = DateSerial(Mid(dateStr, 1, 4), Mid(dateStr, 5, 2), Mid(dateStr, 7, 2))
'        prof = Trim(Mid(components(i), 9, 4))
'        pArr = CInt(Mid(components(i), 13, 4))
'        wsSommaire.Cells(r2, 1).value = Format$(dateStr, "yyyy-mm-dd")
'        wsSommaire.Cells(r2, 2).value = prof
'        wsSommaire.Cells(r2, 3).value = arrHres(pArr, 1)                    'Hres inscrites
'        wsSommaire.Cells(r2, 4).value = arrHres(pArr, 2)                    'Hres d�truites
'        wsSommaire.Cells(r2, 5).value = arrHres(pArr, 1) - arrHres(pArr, 2) 'Hres Nettes
'        wsSommaire.Cells(r2, 6).value = arrHres(pArr, 3)                    'Hres Facturables
'        wsSommaire.Cells(r2, 7).value = arrHres(pArr, 4)                    'Hres Non/Facturables
'        wsSommaire.Cells(r2, 8).value = arrHres(pArr, 5)                    'Hres Factur�es
'        wsSommaire.Cells(r2, 9).value = arrHres(pArr, 6)                    'Hres TEC
'        r2 = r2 + 1
'    Next i
    
    'Ajustement des formats
'    wsSommaire.Range("A2:A" & r2 - 1).NumberFormat = "yyyy-MM-dd"
'    wsSommaire.Range("C2:I" & r2 - 1).NumberFormat = "#,##0.00"
'    wsSommaire.Range("C2:I" & r2 - 1).HorizontalAlignment = xlRight
'    wsSommaire.columns("C:I").ColumnWidth = 10
    
Clean_Exit:

    'Cleaning memory - 2024-09-05 @ 05:44
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

Sub ADMIN_DataFiles_Folder_Selection() '2024-03-28 @ 14:10

    Dim SharedFolder As FileDialog: Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With SharedFolder
        .Title = "Choisir le r�pertoire de donn�es partag�es, selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F5").value = .selectedItems(1)
        End If
    End With
    
    'Clean up
    Set SharedFolder = Nothing
    
End Sub

Sub ADMIN_Invoices_Excel_Folder_Selection() '2024-08-04 @ 07:30

    Dim SharedFolder As FileDialog: Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With SharedFolder
        .Title = "Choisir le r�pertoire des factures (Format Excel)"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F7").value = .selectedItems(1)
        End If
    End With
    
    'Clean up
    Set SharedFolder = Nothing
    
End Sub

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
    wsName = r.Worksheet.name
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(wsName)
    ws.columns.AutoFit
    
    'Clean up
    Set r = Nothing
    Set ws = Nothing

End Sub

Sub Add_Message_To_WorkSheet(ws As Worksheet, r As Long, c As Long, m As String)

    ws.Cells(r, c).value = m
    If c = 1 Then
        ws.Cells(r, c).Font.Bold = True
    End If

End Sub
Sub ADMIN_PDF_Folder_Selection() '2024-03-28 @ 14:10

    Dim PDFFolder As FileDialog: Set PDFFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With PDFFolder
        .Title = "Choisir le r�pertoire des copies de facture (PDF), selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F6").value = .selectedItems(1)
        End If
    End With
    
    'Clean up
    Set PDFFolder = Nothing

End Sub

Sub Apply_Conditional_Formatting_Alternate(rng As Range, headerRows As Long, Optional EmptyLine As Boolean = False)

    'Avons-nous un Range valide ?
    If rng Is Nothing Or rng.rows.count <= headerRows Then
        Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = rng.Worksheet
    Dim dataRange As Range
    
   ' D�finir la plage de donn�es � laquelle appliquer la mise en forme conditionnelle, en
    'excluant les lignes d'en-t�te
    Set dataRange = rng.Resize(rng.rows.count - headerRows).Offset(headerRows, 0)
    
    'Effacer les formats conditionnels existants sur la plage de donn�es
    dataRange.Interior.ColorIndex = xlNone

    'Appliquer les couleurs en alternance
    Dim i As Long
    For i = 1 To dataRange.rows.count
        'V�rifier la position r�elle de la ligne dans la feuille
        If (dataRange.rows(i).Row + headerRows) Mod 2 = 0 Then
            dataRange.rows(i).Interior.Color = RGB(173, 216, 230) ' Bleu p�le
        End If
    Next i
    
    'Clean up
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
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy/mm/dd"
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
                .Range("A2:Q" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy/mm/dd"
                .Range("C2:C" & lastUsedRow & ", " & _
                       "D2:D" & lastUsedRow & ", " & _
                       "F2:F" & lastUsedRow & ", " & _
                       "H2:H" & lastUsedRow & ", " & _
                       "P2:P" & lastUsedRow).HorizontalAlignment = xlLeft
                With .Range("J2:O" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
                End With
                .Range("A1").CurrentRegion.EntireColumn.AutoFit
            End With
        
        Case "wshENC_D�tails"
            With wshENC_D�tails
                .Range("A2:A" & lastUsedRow & ", B2:B" & lastUsedRow & ", D2:D" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("C2:C" & lastUsedRow & ", E2:EB" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("D2:D" & lastUsedRow).NumberFormat = "yyyy/mm/dd"
                .Range("E2:E" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("E2:E" & lastUsedRow).NumberFormat = "#,##0.00"
            End With
        
        Case "wshENC_Ent�te"
            With wshENC_Ent�te
                .Range("A2:A" & lastUsedRow & ", B2:B" & lastUsedRow & ", D2:D" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy/mm/dd"
                .Range("C2:C" & lastUsedRow & ", E2:E" & lastUsedRow & ", G2:G" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("F2:F" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("F2:F" & lastUsedRow).NumberFormat = "#,##0.00 $"
            End With
        
        Case "wshFAC_Comptes_Clients"
            With wshFAC_Comptes_Clients
                .Range("A2:B" & lastUsedRow & ", " & _
                       "D2:G" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("C2:C" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("H2:J" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy/mm/dd"
                .Range("G2:G" & lastUsedRow).NumberFormat = "yyyy/mm/dd"
                .Range("H2:J" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("A1").CurrentRegion.EntireColumn.AutoFit
            End With
        
        Case "wshFAC_D�tails"
            With usedRange
                .Range("A2:A" & lastUsedRow & ", C2:C" & lastUsedRow & ", F2:F" & lastUsedRow & ", G2:G" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("D2:E" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("C2:C" & lastUsedRow).NumberFormat = "#,##0.00"
                .Range("D2:E" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("H2:H" & lastUsedRow & ", J2:J" & lastUsedRow & ", L2:L" & lastUsedRow & ", N2:T" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("O2:O" & lastUsedRow & ", Q2:Q" & lastUsedRow).NumberFormat = "#0.000 %"
            End With
        
        Case "wshFAC_Ent�te"
            With wshFAC_Ent�te
                .Range("A2:D" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy/mm/dd"
                .Range("E2:I" & lastUsedRow & ", K2:K" & lastUsedRow & ", M2:M" & lastUsedRow & ", O2:O" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("J2:J" & lastUsedRow & ", L2:L" & lastUsedRow & ", N2:N" & lastUsedRow & ", P2:V" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("J2:J" & lastUsedRow & ", L2:L" & lastUsedRow & ", N2:N" & lastUsedRow & ", P2:V" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("Q2:Q" & lastUsedRow & ",S2:S" & lastUsedRow).NumberFormat = "#0.000 %"
            End With

        Case "wshFAC_Projets_D�tails"
            With wshFAC_Projets_D�tails
                .Range("A2:A" & lastUsedRow & ", C2:G" & lastUsedRow & ", I2:J" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("F2:F" & lastUsedRow).NumberFormat = "yyyy/mm/dd"
                .Range("H2:I" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("H2:H" & lastUsedRow).NumberFormat = "#,##0.00"
                .Range("I2:I" & lastUsedRow).HorizontalAlignment = xlCenter
            End With
        
        Case "wshFAC_Projets_Ent�te"
            With wshFAC_Projets_Ent�te
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
                .Range("B2:B" & lastUsedRow).NumberFormat = "yyyy/mm/dd"
                .Range("C2:C" & lastUsedRow & _
                    ", D2:D" & lastUsedRow & _
                    ", F2:F" & lastUsedRow & _
                    ", I2:I" & lastUsedRow) _
                        .HorizontalAlignment = xlLeft
                With .Range("G2:H" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
                End With
                With .Range("A2:A" & lastUsedRow) _
                    .Range("J2:J" & lastUsedRow).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent5
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            End With
        
        Case "wshTEC_Local"
            With wshTEC_Local
                .Range("A2:P" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("F2:F" & lastUsedRow & ", G2:G" & lastUsedRow & ", I2:I" & lastUsedRow & _
                            ", O2:O" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("H2:H" & lastUsedRow).NumberFormat = "#0.00"
                .Range("K2:K" & lastUsedRow).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                .columns("F").ColumnWidth = 45
                .columns("G").ColumnWidth = 65
                .columns("I").ColumnWidth = 25
            End With

    End Select

    'Clean up
    Set usedRange = Nothing

End Sub

Sub Compare_2_Workbooks_Column_Formatting()                      '2024-08-19 @ 16:24

    'Erase and create a new worksheet for differences
    Dim wsDiff As Worksheet
    Call CreateOrReplaceWorksheet("Diff�rences_Colonnes")
    Set wsDiff = ThisWorkbook.Worksheets("Diff�rences_Colonnes")
    wsDiff.Range("A1").value = "Worksheet"
    wsDiff.Range("B1").value = "Nb. colonnes"
    wsDiff.Range("C1").value = "Colonne"
    wsDiff.Range("D1").value = "Valeur originale"
    wsDiff.Range("E1").value = "Nouvelle valeur"
    Call Make_It_As_Header(wsDiff.Range("A1:E1"))

    'Set your workbooks and worksheets here
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALIT�\GCF_DataFiles\GCF_BD_MASTER_COPY.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER.xlsx")
    
    Dim wso As Worksheet
    Dim wsn As Worksheet
    
    'Loop through each column (assuming both sheets have the same structure)
    Dim col1 As Range, col2 As Range
    Dim diffLog As String
    Dim diffRow As Long, readColumns As Long
    Dim wsName As String
    diffRow = 1
    For Each wso In wb1.Worksheets
        wsName = wso.name
        Set wsn = wb2.Sheets(wsName)
        
        Dim nbCol As Integer
        nbCol = 1
        Do
            nbCol = nbCol + 1
        Loop Until wso.Cells(1, nbCol).value = ""
        nbCol = nbCol - 1
        
        diffRow = diffRow + 1
        wsDiff.Cells(diffRow, 1).value = wsName
        wsDiff.Cells(diffRow, 2).value = nbCol
        
        Dim i As Integer
        For i = 1 To nbCol
            Set col1 = wso.columns(i)
            Set col2 = wsn.columns(i)
            readColumns = readColumns + 1
            
            'Compare Font Name
            If col1.Font.name <> col2.Font.name Then
                diffLog = diffLog & "Column " & i & " Font Name differs: " & col1.Font.name & " vs " & col2.Font.name & vbCrLf
                wsDiff.Cells(diffRow, 3).value = i
                wsDiff.Cells(diffRow, 4).value = col1.Font.name
                wsDiff.Cells(diffRow, 5).value = col2.Font.name
            End If
            
            'Compare Font Size
            If col1.Font.size <> col2.Font.size Then
                diffLog = diffLog & "Column " & i & " Font Size differs: " & col1.Font.size & " vs " & col2.Font.size & vbCrLf
                wsDiff.Cells(diffRow, 3).value = i
                wsDiff.Cells(diffRow, 4).value = col1.Font.size
                wsDiff.Cells(diffRow, 5).value = col2.Font.size
            End If
            
            'Compare Column Width
            If col1.ColumnWidth <> col2.ColumnWidth Then
                diffLog = diffLog & "Column " & i & " Width differs: " & col1.ColumnWidth & " vs " & col2.ColumnWidth & vbCrLf
                wsDiff.Cells(diffRow, 3).value = i
                wsDiff.Cells(diffRow, 4).value = col1.ColumnWidth
                wsDiff.Cells(diffRow, 5).value = col2.ColumnWidth
            End If
            
            'Compare Number Format
            If col1.NumberFormat <> col2.NumberFormat Then
                diffLog = diffLog & "Column " & i & " Number Format differs: " & col1.NumberFormat & " vs " & col2.NumberFormat & vbCrLf
                wsDiff.Cells(diffRow, 3).value = i
                wsDiff.Cells(diffRow, 4).value = col1.NumberFormat
                wsDiff.Cells(diffRow, 5).value = col2.NumberFormat
            End If
            
            'Compare Horizontal Alignment
            If col1.HorizontalAlignment <> col2.HorizontalAlignment Then
                diffLog = diffLog & "Column " & i & " Horizontal Alignment differs: " & col1.HorizontalAlignment & " vs " & col2.HorizontalAlignment & vbCrLf
                wsDiff.Cells(diffRow, 3).value = i
                wsDiff.Cells(diffRow, 4).value = col1.HorizontalAlignment
                wsDiff.Cells(diffRow, 5).value = col2.HorizontalAlignment
            End If
    
            'Compare Background Color
            If col1.Interior.Color <> col2.Interior.Color Then
                diffLog = diffLog & "Column " & i & " Background Color differs: " & col1.Interior.Color & " vs " & col2.Interior.Color & vbCrLf
                wsDiff.Cells(diffRow, 3).value = i
                wsDiff.Cells(diffRow, 4).value = col1.Interior.Color
                wsDiff.Cells(diffRow, 5).value = col2.Interior.Color
            End If
    
        Next i
        
    Next wso
    
    wsDiff.columns.AutoFit
    wsDiff.Range("B:E").columns.HorizontalAlignment = xlCenter
    
    'Result print setup - 2024-08-05 @ 05:16
    diffRow = diffRow + 2
    wsDiff.Range("A" & diffRow).value = "**** " & Format$(readColumns, "###,##0") & _
                                        " colonnes analys�es dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:E" & diffRow)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:E" & diffRow)
    Dim header1 As String: header1 = wb1.name & " vs. " & wb2.name
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsDiff, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    'Output differences
    If diffLog <> "" Then
        MsgBox "Diff�rences trouv�es:" & vbCrLf & diffLog
    Else
        MsgBox "Aucune diff�rence dans les colonnes."
    End If
    
    'Clean up
    Set col1 = Nothing
    Set col2 = Nothing
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set wb1 = Nothing
    Set wb2 = Nothing
    Set wsn = Nothing
    Set wso = Nothing
    Set wsDiff = Nothing
    
End Sub

Sub Compare_2_Workbooks_Cells_Level()                      '2024-08-20 @ 05:14

    'Erase and create a new worksheet for differences
    Dim wsDiff As Worksheet
    Call CreateOrReplaceWorksheet("Diff�rences_Lignes")
    Set wsDiff = ThisWorkbook.Worksheets("Diff�rences_Lignes")
    wsDiff.Range("A1").value = "Worksheet"
    wsDiff.Range("B1").value = "Prod_Cols"
    wsDiff.Range("C1").value = "Dev_Cols"
    wsDiff.Range("D1").value = "Prod_Rows"
    wsDiff.Range("E1").value = "Dev_Rows"
    wsDiff.Range("F1").value = "Ligne #"
    wsDiff.Range("G1").value = "Colonne"
    wsDiff.Range("H1").value = "Prod_Value"
    wsDiff.Range("I1").value = "Dev_Value"
    Call Make_It_As_Header(wsDiff.Range("A1:I1"))

    'Set your workbooks and worksheets here
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALIT�\GCF_DataFiles\GCF_BD_MASTER_COPY.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER.xlsx")
    
    Dim diffRow As Long
    diffRow = 1
    diffRow = diffRow + 1
    wsDiff.Cells(diffRow, 1).value = "Prod: " & wb1.name
    diffRow = diffRow + 1
    wsDiff.Cells(diffRow, 1).value = "Dev : " & wb2.name
    
    Dim wsProd As Worksheet
    Dim wsDev As Worksheet
    
    'Loop through each column (assuming both sheets have the same structure)
    Dim diffLogMess As String
    Dim readRows As Long
    Dim wsName As String
    For Each wsProd In wb1.Worksheets
        wsName = wsProd.name
        Set wsDev = wb2.Sheets(wsName)
        
        'Determine number of columns and rows in Prod Workbook
        Dim arr(1 To 30) As String
        Dim nbColProd As Integer, nbRowProd As Long
        nbColProd = 0
        Do
            nbColProd = nbColProd + 1
            arr(nbColProd) = wsProd.Cells(1, nbColProd).value
            Debug.Print wsProd.name, " Prod: ", wsProd.Cells(1, nbColProd).value
        Loop Until wsProd.Cells(1, nbColProd).value = ""
        nbColProd = nbColProd - 1
        nbRowProd = wsProd.Cells(wsProd.rows.count, "A").End(xlUp).Row
        
        'Determine number of columns and rows in Dev Workbook
        Dim nbColDev As Integer, nbRowDev As Long
        nbColDev = 0
        Do
            nbColDev = nbColDev + 1
            Debug.Print wsDev.name, " Dev : ", wsDev.Cells(1, nbColDev).value
        Loop Until wsProd.Cells(1, nbColDev).value = ""
        nbColDev = nbColDev - 1
        nbRowDev = wsDev.Cells(wsDev.rows.count, "A").End(xlUp).Row
        
        diffRow = diffRow + 2
        wsDiff.Cells(diffRow, 1).value = wsName
        wsDiff.Cells(diffRow, 2).value = nbColProd
        wsDiff.Cells(diffRow, 3).value = nbColDev
        wsDiff.Cells(diffRow, 4).value = nbRowProd
        wsDiff.Cells(diffRow, 5).value = nbRowDev
        
        Dim nbRow As Long
        If nbRowProd > nbRowDev Then
            wsDiff.Cells(diffRow, 6).value = "Le client a ajout� " & nbRowProd - nbRowDev & " lignes dans la feuille"
            nbRow = nbRowProd
        End If
        If nbRowProd < nbRowDev Then
            wsDiff.Cells(diffRow, 6).value = "Le dev a ajout� " & nbRowDev - nbRowProd & " lignes dans la feuille"
            nbRow = nbRowDev
        End If
        
        Dim rowProd As Range, rowDev As Range
        Dim i As Long, prevI As Long, j As Integer
        For i = 1 To nbRow
            Set rowProd = wsProd.rows(i)
            Set rowDev = wsDev.rows(i)
            readRows = readRows + 1
            
            For j = 1 To nbColProd
                If wsProd.rows.Cells(i, j).value <> wsDev.rows.Cells(i, j).value Then
                    diffLogMess = diffLogMess & "Cell(" & i & "," & j & ") was '" & _
                                  wsProd.rows.Cells(i, j).value & "' is now '" & _
                                  wsDev.rows.Cells(i, j).value & "'" & vbCrLf
                    diffRow = diffRow + 1
                    If i <> prevI Then
                        wsDiff.Cells(diffRow, 6).value = "Ligne # " & i
                        prevI = i
                    End If
                    wsDiff.Cells(diffRow, 7).value = j & "-" & arr(j)
                    wsDiff.Cells(diffRow, 8).value = wsProd.rows.Cells(i, j).value
                    wsDiff.Cells(diffRow, 9).value = wsDev.rows.Cells(i, j).value
                End If
            Next j
            
        Next i
        
    Next wsProd
    
    wsDiff.columns.AutoFit
    wsDiff.Range("B:E").columns.HorizontalAlignment = xlCenter
    wsDiff.Range("F:I").columns.HorizontalAlignment = xlLeft
    
    'Result print setup - 2024-08-20 @ 05:48
    diffRow = diffRow + 2
    wsDiff.Range("A" & diffRow).value = "**** " & Format$(readRows, "###,##0") & _
                                        " lignes analys�es dans l'ensemble du Workbook ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:I" & diffRow)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:I" & diffRow)
    Dim header1 As String: header1 = wb1.name & " vs. " & wb2.name
    Dim header2 As String: header2 = "Changements de lignes ou cellules"
    Call Simple_Print_Setup(wsDiff, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    'Output differences
    If diffLogMess <> "" Then
        MsgBox "Diff�rences trouv�es:" & vbCrLf & diffLogMess
    Else
        MsgBox "Aucune diff�rence dans les lignes."
    End If
    
    'Clean up
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set rowDev = Nothing
    Set rowProd = Nothing
    Set wb1 = Nothing
    Set wb2 = Nothing
    Set wsDev = Nothing
    Set wsProd = Nothing
    Set wsDiff = Nothing
    
End Sub

Sub Fix_Font_Size_And_Family(r As Range, ff As String, fs As Long)

    'r is the range
    'ff is the Font Family
    'fs is the Font Size
    
    With r.Font
        .name = ff
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
    strOutput = "X_TEC_D�placements"
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
    lastUsedRowClientMF = wsMF.Cells(wsMF.rows.count, "A").End(xlUp).Row
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
    lastUsedRowTEC = wsTEC.Cells(wsTEC.rows.count, "A").End(xlUp).Row
    
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
                    wsOutput.Cells(rowOutput, 3).value = clientData(1, 1)
                    wsOutput.Cells(rowOutput, 5).value = clientData(1, 6)
                    wsOutput.Cells(rowOutput, 6).value = clientData(1, 7)
                    wsOutput.Cells(rowOutput, 7).value = clientData(1, 8)
                    wsOutput.Cells(rowOutput, 8).value = clientData(1, 9)
                    wsOutput.Cells(rowOutput, 9).value = clientData(1, 10)
                End If
                rowOutput = rowOutput + 1
        End If
    Next i
    
    'Colonne des Heures
    wsOutput.Range("D2:D" & rowOutput - 1).NumberFormat = "##0.00"
    
    'Tri des donn�es
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

    'Am�liore le Look (saute 1 ligne entre chaque jour)
    For i = rowOutput To 3 Step -1
        If Len(Trim(wsOutput.Cells(i, 3).value)) > 0 Then
            If wsOutput.Cells(i, 2).value <> wsOutput.Cells(i - 1, 2).value Then
                wsOutput.rows(i).Insert Shift:=xlDown
                wsOutput.Cells(i, 1).value = wsOutput.Cells(i - 1, 2).value
            End If
        End If
    Next i
    
    rowOutput = wsOutput.Cells(wsOutput.rows.count, "A").End(xlUp).Row
    
    'Am�liore le Look (cache la date, le client et l'adresse si deux charges & +)
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
    rowOutput = wsOutput.Cells(wsOutput.rows.count, "A").End(xlUp).Row
    
    For i = 3 To rowOutput
        If wsOutput.Cells(i, 1).value > wsOutput.Cells(i - 1, 1).value Then
            wsOutput.Cells(i, 2).Font.Bold = True
        Else
            wsOutput.Cells(i, 2).value = ""
        End If
    Next i
    rowOutput = rowOutput + 2
    wsOutput.Range("A" & rowOutput).value = "**** " & Format$(lastUsedRowTEC - 2, "###,##0") & _
                                        " charges de temps analys�es dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("B2:K" & rowOutput)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
'    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:I" & rowOutput)
    Dim header1 As String: header1 = "Liste des TEC pour Guillaume"
    Dim header2 As String: header2 = "P�riode du " & dateFrom & " au " & dateTo
    Call Simple_Print_Setup(wsOutput, rngArea, header1, header2, "$1:$1", "P")
    
    'Clean up
    Set rngArea = Nothing
    Set rngClientsMF = Nothing
    Set wsOutput = Nothing
    Set wsMF = Nothing
    Set wsTEC = Nothing
    
End Sub

Sub Get_Date_Derniere_Modification(fileName As String, ByRef ddm As Date, _
                                    ByRef jours As Long, ByRef heures As Long, _
                                    ByRef minutes As Long, ByRef secondes As Long)
    
    'Cr�er une instance de FileSystemObject
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'Obtenir le fichier
    Dim fichier As Object: Set fichier = FSO.GetFile(fileName)
    
    'R�cup�rer la date et l'heure de la derni�re modification
    ddm = fichier.DateLastModified
    
    'Calculer la diff�rence (jours) entre maintenant et la date de la derni�re modification
    Dim diff As Double
    diff = Now - ddm
    
    'Convertir la diff�rence en jours, heures, minutes et secondes
    jours = Int(diff)
    heures = Int((diff - jours) * 24)
    minutes = Int(((diff - jours) * 24 - heures) * 60)
    secondes = Int(((((diff - jours) * 24 - heures) * 60) - minutes) * 60)
    
    ' Lib�rer les objets
    Set fichier = Nothing
    Set FSO = Nothing
    
End Sub

Sub LireFichierLogSaisieHeuresTXT() '2024-10-17 @ 20:13
    
    'Initialisation de la bo�te de dialogue FileDialog pour choisir le fichier
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Configuration des filtres de fichiers (TXT uniquement)
    fd.Title = "S�lectionnez un fichier TXT"
    fd.Filters.Clear
    fd.Filters.Add "Fichiers Texte", "*.txt"
    
    'Si l'utilisateur s�lectionne un fichier, filePath contiendra son chemin
    Dim filePath As String
    If fd.show = -1 Then
        filePath = fd.selectedItems(1)
    Else
        MsgBox "Aucun fichier s�lectionn�.", vbExclamation
        Exit Sub
    End If
    
    'Ouvre le fichier en mode lecture
    Dim FileNum As Integer
    FileNum = FreeFile
    Open filePath For Input As FileNum
    
    'Initialise la ligne de d�part pour ins�rer les donn�es dans Excel
    Dim ligneNum As Long
    ligneNum = 1
    
    'Lire chaque ligne du fichier
    Dim ligne As String
    Dim champs() As String
    Dim j As Long

    Do While Not EOF(FileNum)
        Line Input #FileNum, ligne
        
        'S�parer les champs par le s�parateur " | "
        champs = Split(ligne, " | ")
        
        'Ins�rer les champs dans les colonnes de la feuille Excel
        For j = LBound(champs) To UBound(champs)
            Cells(ligneNum, j + 1).value = champs(j)
        Next j
        
        'Passer � la ligne suivante
        ligneNum = ligneNum + 1
    Loop
    
    'Fermer le fichier
    Close FileNum
    
    'Clean up
    Set fd = Nothing
    
    MsgBox "Le fichier a �t� import� avec succ�s.", vbInformation
    
End Sub

Sub CorrigerDatesAvecHeures_ColonnesSpecifiques()
    
    'Initialisation de la bo�te de dialogue FileDialog pour choisir le fichier Excel
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Configuration des filtres de fichiers (Excel uniquement)
    fd.Title = "S�lectionnez un fichier Excel"
    fd.Filters.Clear
    fd.Filters.Add "Fichiers Excel", "*.xlsx; *.xlsm"
    
    'Si l'utilisateur s�lectionne un fichier, filePath contiendra son chemin
    Dim filePath As String
    Dim fileSelected As Boolean
    If fd.show = -1 Then
        filePath = fd.selectedItems(1)
        fileSelected = True
    Else
        MsgBox "Aucun fichier s�lectionn�.", vbExclamation
        fileSelected = False
    End If
    
    'Ouvrir le fichier s�lectionn� s'il y en a un
    Dim wb As Workbook
    If fileSelected Then
        Set wb = Workbooks.Open(filePath)
        
        'D�finir les colonnes sp�cifiques � nettoyer pour chaque feuille
        Dim colonnesANettoyer As Dictionary
        Set colonnesANettoyer = CreateObject("Scripting.Dictionary")
        
        'Ajouter des feuilles et colonnes sp�cifiques (exemple)
        colonnesANettoyer.Add "DEB_Recurrent", Array("B") 'V�rifier la colonne B
        colonnesANettoyer.Add "DEB_Trans", Array("B") 'V�rifier la colonne B
        
        colonnesANettoyer.Add "ENC_D�tails", Array("D") 'V�rifier la colonne D
        colonnesANettoyer.Add "ENC_Ent�te", Array("B") 'V�rifier la colonne B
        
        colonnesANettoyer.Add "FAC_Comptes_Clients", Array("B", "G") 'V�rifier et corriger les colonnes B & G
        colonnesANettoyer.Add "FAC_Ent�te", Array("B") 'V�rifier et corriger la colonne B
        colonnesANettoyer.Add "FAC_Projets_D�tails", Array("F") 'V�rifier et corriger la colonne F
        colonnesANettoyer.Add "FAC_Projets_Ent�te", Array("D") 'V�rifier et corriger la colonne D

        colonnesANettoyer.Add "GL_Trans", Array("B") 'V�rifier et corriger la colonne B

        colonnesANettoyer.Add "TEC_Local", Array("D") 'V�rifier et corriger la colonne D
        
        'Parcourir chaque feuille d�finie dans le dictionnaire
        Dim ws As Worksheet
        Dim cell As Range
        Dim dateOnly As Date
        Dim wsName As Variant
        Dim cols As Variant
        Dim col As Variant
        
        For Each wsName In colonnesANettoyer.keys
            'V�rifier si la feuille existe dans le classeur
            On Error Resume Next
            Set ws = wb.Sheets(wsName)
            Debug.Print wsName
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                'R�cup�rer les colonnes � traiter pour cette feuille
                cols = colonnesANettoyer(wsName)
                
                'Parcourir chaque colonne sp�cifi�e
                For Each col In cols
                    'Parcourir chaque cellule de la colonne sp�cifi�e
                    For Each cell In ws.columns(col).SpecialCells(xlCellTypeConstants)
                        'V�rifier si la cellule contient une date avec une heure
                        If IsDate(cell.value) Then
                            'V�rifier si la valeur contient des heures (fraction d�cimale)
                            If cell.value <> Int(cell.value) Then
                                'Garde uniquement la partie date (sans heure)
                                Debug.Print "", wsName & " - " & col & " - " & cell.value
                                dateOnly = Int(cell.value)
                                cell.value = dateOnly
                            End If
                        End If
                    Next cell
                Next col
            End If
        Next wsName
        
        'Sauvegarder les modifications
        wb.Save
        wb.Close
        
    End If
    
    'Clean up
    Set cell = Nothing
    Set col = Nothing
    Set colonnesANettoyer = Nothing
    Set fd = Nothing
    Set wb = Nothing
    Set ws = Nothing
    Set wsName = Nothing
    
    MsgBox "Les dates ont �t� corrig�es pour les colonnes sp�cifiques.", vbInformation

End Sub

Sub Search_Unclean_Set()

    Dim ws As Worksheet: Set ws = Feuil4
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "B").End(xlUp).Row
    
    Dim strSet As String
    Dim strForEach As String
    Dim strNothing As String
    Dim code As String
    Dim saveModule As String
    Dim saveLineNo As String
    Dim saveProcedure As String
    Dim wsOutput As Worksheet: Set wsOutput = Feuil3
    
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
        'On change de proc�dure
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
'                            Debug.Print i & " - " & saveModule & ":" & saveProcedure & "." & saveLineNo & " - arrSet(" & j & ") = " & arrSet(j) & " n'existe pas dans '" & strNothing & "'"
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
'                            Debug.Print i & " - " & saveModule & ":" & saveProcedure & "." & saveLineNo & " - arrForEach(" & j & ") = " & arrForEach(j) & " n'existe pas dans '" & strNothing & "'"
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
    
    MsgBox "Traiement termin� " & i
    
End Sub
