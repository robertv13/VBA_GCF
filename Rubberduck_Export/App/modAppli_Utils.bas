Attribute VB_Name = "modAppli_Utils"
Option Explicit

Sub Clone_Last_Line_Formatting_For_New_Records(workbookPath As String, wSheet As String, numberRows As Long)

    'Open the workbook
    Dim wb As Workbook: Set wb = Workbooks.Open(workbookPath)
    Dim ws As Worksheet: Set ws = wb.Sheets(wSheet)

    'Find the last row with data in column A
    Dim lastRow As Long
    lastRow = ws.Range("A9999").End(xlUp).Row
    Dim firstNewRow As Long
    firstNewRow = lastRow - numberRows + 1

    'Set the range for new rows
    Dim newRows As Range
    Set newRows = ws.Range(ws.Cells(firstNewRow, 1), ws.Cells(lastRow, ws.columns.count))

    'Copy formatting from the row above the first new row to the new rows
    ws.rows(firstNewRow - 1).Copy
    newRows.PasteSpecial Paste:=xlPasteFormats

    'Clear the clipboard to avoid Excel's cut-copy mode
    Application.CutCopyMode = False

    'Save and close the workbook
    wb.Close SaveChanges:=True

End Sub

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
                       header2 As String, Optional Orient As String = "L")
    
    On Error GoTo CleanUp
    
    Application.PrintCommunication = False
    
    With ws.PageSetup
        .PrintArea = rng.Address
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = ""
        
        .CenterHeader = "&""-,Gras""&12&K0070C0" & header1 & Chr(10) & "&11" & header2
        
        .LeftFooter = "&11&D - &T"
        .CenterFooter = "&11&KFF0000&A"
        .RightFooter = "Page &P of &N"
        
        .TopMargin = Application.InchesToPoints(0.75)
        .LeftMargin = Application.InchesToPoints(0.15)
        .RightMargin = Application.InchesToPoints(0.15)
        .BottomMargin = Application.InchesToPoints(0.55)
        
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
    If Err.Number <> 0 Then
        MsgBox "Error setting PrintCommunication to True: " & Err.Description, vbCritical
    End If
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
    Set wsTgt = ThisWorkbook.Worksheets("HresJourProf")
    
    Dim lastUsedRowSrc As Long
    lastUsedRowSrc = wsSrc.Cells(wsSrc.rows.count, "A").End(xlUp).Row '2024-08-15 @ 06:17
    
    wsTgt.Range("A2:H" & wsTgt.Cells(wsTgt.rows.count, "A").End(xlUp).Row).ClearContents
    
    'Copy columns A to H (from Source to Target), using Copy and Paste Special
    wsSrc.Range("A2:H" & lastUsedRowSrc).Copy
    wsTgt.Cells(2, 1).PasteSpecial Paste:=xlPasteValues
    
    'Clear the clipboard
    Application.CutCopyMode = False
    
'    Dim i As Long, j As Long
'    For i = 2 To lastUsedRowSrc
'        For j = 1 To 8
'            wsTgt.Cells(i, j).value = wsSrc.Cells(i, j).value
'        Next j
'    Next i

    Call Update_Pivot_Table
    
    MsgBox "L'importation des Heures par Jour / Professionnel est complétée" & _
            vbNewLine & vbNewLine & "Ainsi que la mise à jour du Pivot Table", _
            vbExclamation
    
End Sub

Sub Update_Pivot_Table() '2024-08-15 @ 06:34

    'Define the worksheet containing the data
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HresJourProf")
    
    'Find the last row of your data
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    
    'Define the new data range
    Dim rngData As Range
    Set rngData = ws.Range("A1:H" & lastUsedRow)
    
    'Update the Pivot Table
    Dim pt As PivotTable
    Set pt = ws.PivotTables("ptHresJourProf")
    pt.ChangePivotCache ThisWorkbook.PivotCaches.Create( _
                        SourceType:=xlDatabase, _
                        SourceData:=rngData)
    
    'Refresh the Pivot Table
    pt.RefreshTable

End Sub

'Function CompareComponents(comp1 As String, comp2 As String) As Long
'
'    'Extract date part from each component
''    Dim date1 As String
''    Dim date2 As String
''
''    date1 = Split(comp1, "_")(0)
''    date2 = Split(comp2, "_")(0)
'
'    'Compare 2 components
'    CompareComponents = StrComp(comp1, comp2, vbTextCompare)
'
'End Function
'
Sub Start_Timer(subName As String) '2024-06-06 @ 10:12

    Dim modeOper As Long
    modeOper = 2
    
    'modeOper = 1 - Dump to immediate Window
    If modeOper = 1 Then
        Dim l As Long: l = Len(subName)
        Debug.Print vbNewLine & String(40 + l, "*") & vbNewLine & _
        Format$(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & "Entering: " & subName & _
            vbNewLine & String(40 + l, "*")
    End If

    'modeOper = 2 - Dump to worksheet
    If modeOper = 2 Then
        With wshzDocLogAppli
            Dim lastUsedRow As Long
            lastUsedRow = .Range("A99999").End(xlUp).Row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Fn_Get_Windows_Username
            .Range("B" & lastUsedRow).value = ThisWorkbook.name
            .Range("C" & lastUsedRow).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("D" & lastUsedRow).value = subName & " - entering"
        End With
    End If

End Sub

Sub End_Timer(subName As String, t As Double)

    Dim modeOper As Long
    modeOper = 2 '2024-03-29 @ 11:37
    
    'Allows message to be used - 2024-06-06 @ 11:05
    If InStr(subName, "message:") = 1 Then
        subName = Right(subName, Len(subName) - 8)
    Else
        subName = subName & " - exiting"
    End If
    
    'modeOper = 1 - Dump to immediate Window
    If modeOper = 1 Then
        Dim l As Long: l = Len(subName)
        Debug.Print vbNewLine & String(40 + l, "*") & vbNewLine & _
        Format$(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & subName & " = " _
        & Format$(Timer - t, "##0.0000") & " secondes" & vbNewLine & String(40 + l, "*")
    End If

    'modeOper = 2 - Dump to worksheet
    If modeOper = 2 Then
        With wshzDocLogAppli
            Dim lastUsedRow As Long
            lastUsedRow = .Range("A99999").End(xlUp).Row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Fn_Get_Windows_Username
            .Range("B" & lastUsedRow).value = ThisWorkbook.name
            .Range("C" & lastUsedRow).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("D" & lastUsedRow).value = subName
            If t Then
                .Range("E" & lastUsedRow).value = Round(Timer - t, 4)
                .Range("E" & lastUsedRow).NumberFormat = "#,##0.0000"
            End If
        End With
    End If

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
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modAppli_Utils:CreateOrReplaceWorksheet()")
    
    Dim ws As Worksheet
    Dim wsExists As Boolean
    wsExists = False
    
    'Check if the worksheet exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = wsName Then
            wsExists = True
            Exit For
        End If
    Next ws
    
    'If the worksheet exists, delete it
    If wsExists Then
        Application.DisplayAlerts = False
        ws.delete
        Application.DisplayAlerts = True
    End If
    
    'Add the new worksheet
    Set ws = ThisWorkbook.Worksheets.add(before:=wshMenu)
    ws.name = wsName

    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
    Call End_Timer("modAppli_Utils:CreateOrReplaceWorksheet()", timerStart)

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
        MsgBox "Il existe des références circulaires dans le Workbook dans les cellules suivantes:" & vbCrLf & circRef, vbExclamation
    Else
        MsgBox "Il n'existe aucune référence circulaire dans ce Workbook .", vbInformation
    End If
    
End Sub

Public Sub Integrity_Verification() '2024-07-06 @ 12:56

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:Integrity_Verification", 0)

    Application.ScreenUpdating = False
    
    Call Erase_And_Create_Worksheet("X_Analyse_Intégrité")
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    wsOutput.Range("A1").value = "Feuille"
    wsOutput.Range("B1").value = "Message"
    wsOutput.Range("C1").value = "TimeStamp"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"))

    Call Erase_And_Create_Worksheet("X_Heures_Jour_Prof")
    Dim wsSommaire As Worksheet: Set wsSommaire = ThisWorkbook.Worksheets("X_Heures_Jour_Prof")
    wsSommaire.Range("A1").value = "Date"
    wsSommaire.Range("B1").value = "Prof."
    wsSommaire.Range("C1").value = "H/Saisies"
    wsSommaire.Range("D1").value = "H/Détruites"
    wsSommaire.Range("E1").value = "H/Fact"
    wsSommaire.Range("F1").value = "H/NFact"
    wsSommaire.Range("G1").value = "H/Facturées"
    wsSommaire.Range("H1").value = "H/TEC"
    Call Make_It_As_Header(wsSommaire.Range("A1:H1"))

    'Data starts at row 2
    Dim r As Long: r = 2
    Dim readRows As Long
    
    'dnrPlanComptable ----------------------------------------------------- Plan Comptable
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Plan Comptable")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    
    Call check_Plan_Comptable(r, readRows)

    'wshBD_Clients --------------------------------------------------------------- Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "BD_Clients")
    
    Call Client_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_Clients(r, readRows)

    'wshBD_Fournisseurs ----------------------------------------------------- Fournisseurs
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "BD_Fournisseurs")
    
    Call Fournisseur_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_Fournisseurs(r, readRows)
    
    'wshENC_Détails ---------------------------------------------------------- ENC_Détails
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "ENC_Détails")
    
    Call ENC_Détails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "ENC_Détails a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_ENC_Détails(r, readRows)
    
    'wshENC_Entête ------------------------------------------------------------ ENC_Entête
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "ENC_Entête")
    
    Call ENC_Entête_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "ENC_Entête a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_ENC_Entête(r, readRows)
    
    'wshFAC_Détails ---------------------------------------------------------- FAC_Détails
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Détails")
    
    Call FAC_Détails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Détails a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_FAC_Détails(r, readRows)
    
    'wshFAC_Entête ------------------------------------------------------------ FAC_Entête
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Entête")
    
    Call FAC_Entête_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Entête a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_FAC_Entête(r, readRows)
    
    'wshFAC_Comptes_Clients ------------------------------------------ FAC_Comptes_Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Comptes_Clients")
    
    Call FAC_Comptes_Clients_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Comptes_Clients a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_FAC_Comptes_Clients(r, readRows)
    
    'wshFAC_Projets_Détails ------------------------------------------ FAC_Projets_Détails
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Projets_Détails")
    
    Call FAC_Projets_Détails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Projets_Détails a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_FAC_Projets_Détails(r, readRows)
    
    'wshFAC_Projets_Entête -------------------------------------------- FAC_Projets_Entête
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Projets_Entête")
    
    Call FAC_Projets_Entête_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Projets_Entête a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_FAC_Projets_Entête(r, readRows)
    
    'wshGL_Trans ---------------------------------------------------------------- GL_Trans
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "GL_Trans")
    
    Call GL_Trans_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "GL_Trans a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))

    Call check_GL_Trans(r, readRows)
    
    'wshTEC_TdB_Data -------------------------------------------------------- TEC_TdB_Data
    
    Call TEC_TdB_Update_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "TEC_TdB_Data")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    
    Call check_TEC_TdB_Data(r, readRows)
    
    'wshTEC_Local -------------------------------------------------------------- TEC_Local
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "TEC_Local")
    
    Call TEC_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "TEC_Local a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "mm/dd/yyyy hh:nn:ss"))
    r = r + 1
    
    Call check_TEC(r, readRows)
    
    'Adjust the Output Worksheet
    With wsOutput.Range("A2:C" & r).Font
        .name = "Courier New"
        .size = 10
    End With
    
    wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit
    wsSommaire.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
   'Result print setup - 2024-07-20 @ 14:31
    Dim lastUsedRow As Long
    lastUsedRow = r + 1
    wsOutput.Range("A" & lastUsedRow).value = "**** " & Format$(readRows, "###,##0") & _
                                    " lignes analysées dans l'ensemble des tables ***"
    
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "Vérification d'intégrité des tables"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "P")
    
    MsgBox "La vérification d'intégrité est terminé" & vbNewLine & vbNewLine & "Voir la feuille 'X_Analyse_Intégrité'", vbInformation
    
    ThisWorkbook.Worksheets("X_Analyse_Intégrité").Activate
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:Integrity_Verification", startTime)

End Sub

Private Sub check_Clients(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_Clients", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
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
            dict_nom_client.add nom, code
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "À la ligne " & i & ", le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        
        If dict_code_client.Exists(code) = False Then
            dict_code_client.add code, nom
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "À la ligne " & i & ", le code '" & code & "' est un doublon pour le client '" & nom & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
        
    Next i
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 1, "##,##0") & " clients ont été analysés!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de nom")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        r = r + 1
    End If
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de code")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
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

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
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
            dict_nom_fournisseur.add nom, code
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        If dict_code_fournisseur.Exists(code) = False Then
            dict_code_fournisseur.add code, nom
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le code '" & code & "' est un doublon pour le nom '" & nom & "'")
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
    Next i
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 1, "#,##0") & " fournisseurs ont été analysés!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de nom")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        r = r + 1
    End If
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de code")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
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

Private Sub check_ENC_Détails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_ENC_Détails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshENC_Détails
    Dim ws As Worksheet: Set ws = wshENC_Détails
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRowDetails As Long
    lastUsedRowDetails = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastUsedRowDetails <= 2 - headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
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
    lastUsedRowEntete = wsEntete.Cells(wsEntete.rows.count, "A").End(xlUp).Row
    Dim rngEntete As Range: Set rngEntete = wsEntete.Range("A2:A" & lastUsedRowEntete)
    Dim strPmtNo As String
    Dim i As Long
    For i = 2 To lastUsedRowEntete
        strPmtNo = strPmtNo & wsEntete.Range("A" & i).value & "|"
    Next i
    
    'FAC_Entête Worksheet
    Dim wsFACEntete As Worksheet: Set wsFACEntete = wshFAC_Entête
    Dim lastUsedRowFacEntete As Long
    lastUsedRowFacEntete = wsFACEntete.Cells(wsFACEntete.rows.count, "A").End(xlUp).Row
    Dim rngFACEntete As Range: Set rngFACEntete = wsFACEntete.Range("A2:A" & lastUsedRowFacEntete)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshENC_Détails'")
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
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le paiement '" & pmtNo & "' à la ligne " & i & " n'existe pas dans ENC_Entête")
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' du paiement '" & pmtNo & "' n'existe pas dans FAC_Entête")
            r = r + 1
        End If
        
        If IsDate(ws.Range("D" & i).value) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La date '" & ws.Range("D" & i).value & "' du paiment '" & pmtNo & "' est INVALIDE '")
            r = r + 1
        End If
        
        If IsNumeric(ws.Range("E" & i).value) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le montant '" & ws.Range("E" & i).value & "' du paiement '" & pmtNo & "' n'est pas numérique")
            r = r + 1
        End If
        totalEncDetails = totalEncDetails + ws.Range("E" & i).value
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Total des encaissements : " & Format$(totalEncDetails, "#,##0.00 $"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(lastUsedRowDetails - 1, "##,##0") & " lignes de transactions ont été analysées")
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
    
    Call Log_Record("modAppli:check_ENC_Détails", startTime)

End Sub

Private Sub check_ENC_Entête(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_ENC_Entête", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'Clients Master File
    Dim wsClients As Worksheet: Set wsClients = wshBD_Clients
    Dim lastUsedRowClient As Long
    lastUsedRowClient = wsClients.Cells(wsClients.rows.count, "B").End(xlUp).Row
    Dim rngClients As Range: Set rngClients = wsClients.Range("B2:B" & lastUsedRowClient)
    
    'wshENC_Entête
    Dim ws As Worksheet: Set ws = wshENC_Entête
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshENC_Entête'")
    r = r + 1
    
    If lastUsedRow = headerRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim arr As Variant
    arr = wshENC_Entête.Range("A1").CurrentRegion.Offset(1, 0) _
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La date de paiement '" & arr(i, 2) & "' du paiement '" & arr(i, 1) & "' n'est pas VALIDE")
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le client '" & codeClient & "' du paiement '" & pmtNo & "' est INVALIDE")
            r = r + 1
        End If
        totals = totals + arr(i, 6)
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Total des encaissements : " & Format$(totals, "#,##0.00 $"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " factures ont été analysées")
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
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Entête
    Dim lastUsedRowEntete As Long
    lastUsedRowEntete = wsMaster.Cells(wsMaster.rows.count, "A").End(xlUp).Row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A" & 1 + headerRow & ":A" & lastUsedRowEntete)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Détails'")
    r = r + 1
    
    'Transfer data from Worksheet into an Array (arr)
    Dim arr As Variant
    arr = wshFAC_Détails.Range("A1").CurrentRegion.Offset(1, 0).value
    
    'Array pointer
    Dim Row As Long: Row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String, oldInv_No As String
    Dim result As Variant
    For i = LBound(arr, 1) + 2 To UBound(arr, 1) - 1 'Two lines of header !
        Inv_No = CStr(arr(i, 1))
        If Inv_No <> oldInv_No Then
             result = Application.WorksheetFunction.XLookup(Inv_No, _
                                                    rngMaster, _
                                                    rngMaster, _
                                                    "Not Found", _
                                                    0, _
                                                    1)

'            result = Application.WorksheetFunction.XLookup(ws.Cells(i, 1), rngMaster, rngMaster, "Not Found", 0, 1)
            oldInv_No = Inv_No
        End If
        If result = "Not Found" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " n'existe pas dans FAC_Entête")
            r = r + 1
        End If
        If IsNumeric(arr(i, 3)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " le nombre d'heures est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " le taux horaire est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " le montant est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 2, "##,##0") & " lignes de transactions ont été analysées")
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
    lastUsedRow = ws.Range("A9999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Entête'")
    r = r + 1
    
    If lastUsedRow = headerRow Then
        r = r + 1
        GoTo Clean_Exit
    End If

    Dim arr As Variant
    arr = wshFAC_Entête.Range("A1").CurrentRegion.Offset(2, 0) _
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
        End If
        If arr(i, 3) <> "C" And arr(i, 3) <> "AC" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le type de facture '" & arr(i, 3) & "' pour la facture '" & Inv_No & "' est INVALIDE")
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Factures CONFIRMÉES (" & nbFactC & " factures)")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 1), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 1), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 1), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 1), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 1), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 1), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Total Fact. : " & _
            Fn_Pad_A_String(Format$(totals(7, 1), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Acompte payé: " & _
            Fn_Pad_A_String(Format$(totals(8, 1), "##,##0.00$"), " ", 12, "L"))
    r = r + 2
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Factures À CONFIRMER (" & nbFactAC & " factures)")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Honoraires  : " & _
            Fn_Pad_A_String(Format$(totals(1, 2), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Divers - 1  : " & _
            Fn_Pad_A_String(Format$(totals(2, 2), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Divers - 2  : " & _
            Fn_Pad_A_String(Format$(totals(3, 2), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Divers - 3  : " & _
            Fn_Pad_A_String(Format$(totals(4, 2), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     TPS         : " & _
            Fn_Pad_A_String(Format$(totals(5, 2), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     TVQ         : " & _
            Fn_Pad_A_String(Format$(totals(6, 2), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Total Fact. : " & _
            Fn_Pad_A_String(Format$(totals(7, 2), "##,##0.00$"), " ", 12, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Acompte payé: " & _
            Fn_Pad_A_String(Format$(totals(8, 2), "##,##0.00$"), " ", 12, "L"))
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
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
    lastUsedRow = ws.Range("A9999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le type de facture '" & invType & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsDate(CDate(arr(i, 2))) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La date de facture '" & arr(i, 2) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If Fn_Validate_Client_Number(CStr(arr(i, 4))) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le client '" & CStr(arr(i, 4)) & "' de la facture '" & Inv_No & "' est INVALIDE '")
            r = r + 1
        End If
        If arr(i, 5) <> "Paid" And arr(i, 5) <> "Unpaid" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le statut '" & arr(i, 5) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsDate(CDate(arr(i, 7))) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La date due '" & arr(i, 7) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le total de la facture '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsNumeric(arr(i, 9)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le montant payé à date '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If IsNumeric(arr(i, 10)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le solde de la facture '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        'PLUG pour s'assurer que le solde impayé est belt et bien aligner sur le total et $ payé à date
        If arr(i, 10) <> arr(i, 8) - arr(i, 9) Then
            arr(i, 10) = arr(i, 8) - arr(i, 9)
        End If
        If IsNumeric(arr(i, 11)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** L'âge (jours) de la facture '" & arr(i, 8) & "' de la facture '" & Inv_No & "' est INVALIDE")
            r = r + 1
        End If
        If arr(i, 10) = 0 And arr(i, 5) = "Unpaid" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le statut '" & arr(i, 5) & "' de la facture '" & Inv_No & "', avec un solde de " & Format$(arr(i, 10), "#,##0.00 $") & " est INVALIDE")
            r = r + 1
        End If
        If arr(i, 10) <> 0 And arr(i, 5) = "Paid" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le statut '" & arr(i, 5) & "' de la facture '" & Inv_No & "', avec un solde de " & Format$(arr(i, 10), "#,##0.00 $") & " est INVALIDE")
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Factures CONFIRMÉES (" & nbFactC & " factures)")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Total des factures        : " & Fn_Pad_A_String(Format$(totals(1, 1), "###,##0.00$"), " ", 13, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Montants encaissés à date : " & Fn_Pad_A_String(Format$(totals(2, 1), "##,##0.00$"), " ", 13, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Solde à recevoir          : " & Fn_Pad_A_String(Format$(totals(3, 1), "##,##0.00$"), " ", 13, "L"))
    r = r + 2
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Factures À CONFIRMER (" & nbFactAC & " factures)")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Total des factures        : " & Fn_Pad_A_String(Format$(totals(1, 2), "###,##0.00$"), " ", 13, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Montants encaissés à date : " & Fn_Pad_A_String(Format$(totals(2, 2), "##,##0.00$"), " ", 13, "L"))
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Solde à recevoir          : " & Fn_Pad_A_String(Format$(totals(3, 2), "##,##0.00$"), " ", 13, "L"))
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

Private Sub check_FAC_Projets_Détails(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Projets_Détails", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshFAC_Projets_Détails
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If

    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Projets_Entête
    lastUsedRow = wsMaster.Cells(wsMaster.rows.count, "A").End(xlUp).Row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A2:A" & lastUsedRow)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Projets_Détails'")
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
    Dim lookUpValue As Long, result As Variant
    For i = LBound(arr, 1) To UBound(arr, 1)
        projetID = arr(i, 1)
        lookUpValue = projetID
        If projetID <> oldProjetID Then
            result = Application.WorksheetFunction.XLookup(lookUpValue, _
                                                           rngMaster, _
                                                           rngMaster, _
                                                           "Not Found", _
                                                           0, _
                                                           1)
            oldProjetID = projetID
        End If
        If result = "Not Found" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " n'existe pas dans FAC_Projets_Entête")
            r = r + 1
        End If
        If IsNumeric(arr(i, 3)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le ClientID est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le TECID est INVALIDE '" & arr(i, 4) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le ProfID est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Heures sont INVALIDES '" & arr(i, 8) & "'")
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes ont été analysées")
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
    
    Call Log_Record("modAppli:check_FAC_Projets_Détails", startTime)

End Sub

Private Sub check_FAC_Projets_Entête(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_FAC_Projets_Entête", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Projets_Entête'")
    r = r + 1
    
    'Establish the number of rows before transferring it to an Array
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.rows.count
    If numRows <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
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
    For i = LBound(arr, 1) To UBound(arr, 1) 'One line of header !
        projetID = arr(i, 1)
        If IsNumeric(arr(i, 3)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le ClientID est INVALIDE '" & arr(i, 3) & "'")
            r = r + 1
        End If
        If IsDate(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 4) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le total des honoraires est INVALIDE '" & arr(i, 5) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 7)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du premier sommaire sont INVALIDES '" & arr(i, 7) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du premier sommaire est INVALIDE '" & arr(i, 8) & "'")
            r = r + 1
        End If
        If IsNumeric(arr(i, 9)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du premier sommaire sont INVALIDES '" & arr(i, 9) & "'")
            r = r + 1
        End If
        If arr(i, 11) <> "" And IsNumeric(arr(i, 11)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du second sommaire sont INVALIDES '" & arr(i, 11) & "'")
            r = r + 1
        End If
        If arr(i, 12) <> "" And IsNumeric(arr(i, 12)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du second sommaire est INVALIDE '" & arr(i, 12) & "'")
            r = r + 1
        End If
        If arr(i, 13) <> "" And IsNumeric(arr(i, 13)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du second sommaire sont INVALIDES '" & arr(i, 13) & "'")
            r = r + 1
        End If
        If arr(i, 15) <> "" And IsNumeric(arr(i, 15)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du troisième sommaire sont INVALIDES '" & arr(i, 15) & "'")
            r = r + 1
        End If
        If arr(i, 16) <> "" And IsNumeric(arr(i, 16)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du troisième sommaire est INVALIDE '" & arr(i, 16) & "'")
            r = r + 1
        End If
        If arr(i, 17) <> "" And IsNumeric(arr(i, 17)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du troisième sommaire sont INVALIDES '" & arr(i, 17) & "'")
            r = r + 1
        End If
        If arr(i, 19) <> "" And IsNumeric(arr(i, 19)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du quatrième sommaire sont INVALIDES '" & arr(i, 19) & "'")
            r = r + 1
        End If
        If arr(i, 20) <> "" And IsNumeric(arr(i, 20)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du quatrième sommaire est INVALIDE '" & arr(i, 20) & "'")
            r = r + 1
        End If
        If arr(i, 21) <> "" And IsNumeric(arr(i, 21)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du quatrième sommaire sont INVALIDES '" & arr(i, 21) & "'")
            r = r + 1
        End If
        If arr(i, 23) <> "" And IsNumeric(arr(i, 23)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du cinquième sommaire sont INVALIDES '" & arr(i, 23) & "'")
            r = r + 1
        End If
        If arr(i, 24) <> "" And IsNumeric(arr(i, 24)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du cinquième sommaire est INVALIDE '" & arr(i, 24) & "'")
            r = r + 1
        End If
        If arr(i, 25) <> "" And IsNumeric(arr(i, 25)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du cinquième sommaire sont INVALIDES '" & arr(i, 25) & "'")
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_FAC_Projets_Entête", startTime)

End Sub

Private Sub check_GL_Trans(ByRef r As Long, ByRef readRows As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_GL_Trans", 0)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshGL_Trans
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
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
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
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
            dict_GL_Entry.add GL_Entry_No, Row
            sum_arr(Row, 1) = GL_Entry_No
            Row = Row + 1
        End If
        If IsDate(arr(i, 2)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** L'écriture #  " & GL_Entry_No & " ' à la ligne " & i & " a une date INVALIDE '" & arr(i, 2) & "'")
            r = r + 1
        End If
        glCode = arr(i, 5)
        If InStr(1, strCodeGL, glCode + "|:|") = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le compte '" & glCode & "' à la ligne " & i & " est INVALIDE '")
            r = r + 1
        End If
        If glCode = "1100" Then
            arTotal = arTotal + arr(i, 7) - arr(i, 8)
        End If
        glDescr = arr(i, 6)
        If InStr(1, strDescGL, glDescr + "|:|") = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La description du compte '" & glDescr & "' à la ligne " & i & " est INVALIDE")
            r = r + 1
        End If
        dt = arr(i, 7)
        If IsNumeric(dt) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le montant du débit '" & dt & "' à la ligne " & i & " n'est pas une valeur numérique")
            r = r + 1
        End If
        ct = arr(i, 8)
        If IsNumeric(ct) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le montant du débit '" & ct & "' à la ligne " & i & " n'est pas une valeur numérique")
            r = r + 1
        End If
        currentRow = dict_GL_Entry(GL_Entry_No)
        sum_arr(currentRow, 2) = sum_arr(currentRow, 2) + dt
        sum_arr(currentRow, 3) = sum_arr(currentRow, 3) + ct
        If arr(i, 10) <> "" Then
            If IsDate(arr(i, 10)) = False Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le TimeStamp '" & arr(i, 10) & "' à la ligne " & i & " n'est pas une date VALIDE")
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Écriture # " & v & " ne balance pas... Dt = " & Format$(dt, "###,###,##0.00") & " et Ct = " & Format$(ct, "###,###,##0.00"))
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
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Un total de " & dict_GL_Entry.count & " écritures ont été analysées")
    r = r + 1
    
    If cas_hors_balance = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Chacune des écritures balancent au niveau de l'écriture")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_hors_balance & " écriture(s) qui ne balance(nt) pas !!!")
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Les totaux des transactions sont:")
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Dt = " & Format$(sum_dt, "###,###,##0.00 $"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Ct = " & Format$(sum_ct, "###,###,##0.00 $"))
    r = r + 1
    
    If sum_dt - sum_ct <> 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Hors-Balance de " & Format$(sum_dt - sum_ct, "###,###,##0.00 $"))
        r = r + 1
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Au Grand Livre, le solde des Comptes-Clients est de : " & Format$(arTotal, "###,###,##0.00 $"))
    r = r + 2
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set planComptable = Nothing
    Set v = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
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
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
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
    
    Dim i As Long, TECID As Long, profID As String, prof As String, dateTEC As Date
    Dim minDate As Date, maxDate As Date
    Dim hres As Double, estFacturable As Boolean
    Dim estFacturee As Boolean, estDetruit As Boolean
    Dim cas_doublon_TECID As Long, cas_date_invalide As Long, cas_doublon_prof As Long, cas_doublon_client As Long
    Dim cas_hres_invalide As Long, cas_estFacturable_invalide As Long, cas_estFacturee_invalide As Long
    Dim cas_estDetruit_invalide As Long
    Dim total_hres_inscrites As Double, total_hres_detruites As Double, total_hres_facturees As Double
    Dim total_hres_facturable As Double, total_hres_TEC As Double, total_hres_non_facturable As Double
    
    minDate = "12/31/2999"
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        TECID = arr(i, 1)
        prof = arr(i, 2)
        dateTEC = arr(i, 3)
        If IsDate(dateTEC) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** TEC_ID =" & TECID & " a une date INVALIDE '" & dateTEC & " !!!")
            r = r + 1
            cas_date_invalide = cas_date_invalide + 1
        Else
            If dateTEC < minDate Then minDate = dateTEC
            If dateTEC > maxDate Then maxDate = dateTEC
        End If
        hres = arr(i, 5)
        If IsNumeric(hres) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            r = r + 1
            cas_hres_invalide = cas_hres_invalide + 1
        End If
        estFacturable = arr(i, 6)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            r = r + 1
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If
        estFacturee = arr(i, 7)
        If InStr("Vrai^Faux^", estFacturee & "^") = 0 Or Len(estFacturee) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            r = r + 1
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        estDetruit = arr(i, 8)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            r = r + 1
            cas_estDetruit_invalide = cas_estDetruit_invalide + 1
        End If
        
        total_hres_inscrites = total_hres_inscrites + hres
        If estDetruit = "Vrai" Then total_hres_detruites = total_hres_detruites + hres
        
        If estDetruit = "Faux" And estFacturable = "Vrai" Then total_hres_facturable = total_hres_facturable + hres
        If estDetruit = "Faux" And estFacturable = "Faux" Then total_hres_non_facturable = total_hres_non_facturable + hres
        If estDetruit = "Faux" And estFacturee = "Vrai" Then total_hres_facturees = total_hres_facturees + hres
        
        'Dictionary
        If dict_TEC_ID.Exists(TECID) = False Then
            dict_TEC_ID.add TECID, 0
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le TEC_ID '" & TECID & "' est un doublon pour la ligne '" & i & "'")
            r = r + 1
            cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        If dict_prof.Exists(prof & "-" & profID) = False Then
            dict_prof.add prof & "-" & profID, 0
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " charges de temps ont été analysées!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
    If cas_doublon_TECID = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de TEC_ID")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_TECID & " cas de doublons pour les TEC_ID")
        r = r + 1
    End If
    
    If cas_date_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune date INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    r = r + 1
    
    If cas_hres_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune heures INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturable_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturable' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturee_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturee' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        r = r + 1
    End If
    
    If cas_estDetruit_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estDetruit' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La somme des heures donne ce résultat:")
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures inscrites       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures détruites       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures restantes       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures facturables     : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures non_facturables : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures Facturées       : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable - total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures TEC             : " & formattedHours)
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
            dict_descr_GL.add descrGL, codeGL
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La description '" & descrGL & "' est un doublon pour le code de G/L '" & codeGL & "'")
            r = r + 1
            cas_doublon_descr = cas_doublon_descr + 1
        End If
        
        If dict_code_GL.Exists(codeGL) = False Then
            dict_code_GL.add codeGL, descrGL
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
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de description")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_descr & " cas de doublons pour les descriptions")
        r = r + 1
    End If
    
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de code de G/L")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_code & " cas de doublons pour les codes de G/L")
        r = r + 1
    End If
    
    If cas_type = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun type de G/L invalide")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_type & " cas de types de G/L invalides")
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

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:check_GL_Trans", 0)
    
    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Analyse_Intégrité")
    Dim wsSommaire As Worksheet: Set wsSommaire = ThisWorkbook.Worksheets("X_Heures_Jour_Prof")
    
    'wshTEC_Local
    Dim ws As Worksheet: Set ws = wshTEC_Local
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
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
    
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Offset(2)
    Dim dict_TEC_ID As New Dictionary
    Dim dict_prof As New Dictionary
    
    Dim i As Long, TECID As Long, profID As String, prof As String, dateTEC As Date, testDate As Boolean
    Dim minDate As Date, maxDate As Date
    Dim d As Integer, m As Integer, y As Integer, p As Integer
    Dim codeClient As String, nomClient As String
    Dim isClientValid As Boolean
    Dim hres As Double, testHres As Boolean, estFacturable As Boolean
    Dim estFacturee As Boolean, estDetruit As Boolean
    Dim cas_doublon_TECID As Long, cas_date_invalide As Long, cas_doublon_prof As Long, cas_doublon_client As Long
    Dim cas_date_future As Long
    Dim cas_hres_invalide As Long, cas_estFacturable_invalide As Long, cas_estFacturee_invalide As Long
    Dim cas_estDetruit_invalide As Long
    Dim total_hres_inscrites As Double, total_hres_detruites As Double, total_hres_facturees As Double
    Dim total_hres_facturable As Double, total_hres_TEC As Double, total_hres_non_facturable As Double
    Dim keyDate As String
    
    minDate = "12/31/2999"
    Dim bigStrDateProf As String
    Dim arrHres(1 To 10000, 1 To 6) As Variant
    Dim arrRow As Integer, pArr As Integer, rArr As Integer
    
    For i = LBound(arr, 1) To UBound(arr, 1) - 2
        TECID = arr(i, 1)
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le code de client '" & codeClient & "' est INVALIDE !!!")
            r = r + 1
        End If
        nomClient = arr(i, 6)
        hres = arr(i, 8)
        testHres = IsNumeric(hres)
        If testHres = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            r = r + 1
            cas_hres_invalide = cas_hres_invalide + 1
        End If
        estFacturable = arr(i, 10)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            r = r + 1
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If
        estFacturee = arr(i, 12)
        If InStr("Vrai^Faux^", estFacturee & "^") = 0 Or Len(estFacturee) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            r = r + 1
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        estDetruit = arr(i, 14)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            r = r + 1
            cas_estDetruit_invalide = cas_estDetruit_invalide + 1
        End If
        
        Dim h(1 To 6) As Double
        h(1) = 0
        total_hres_inscrites = total_hres_inscrites + hres
        h(1) = hres
        
        h(2) = 0
        If estDetruit = "Vrai" Then
            total_hres_detruites = total_hres_detruites + hres
            h(2) = hres
        End If
        
        h(3) = 0
        If estDetruit = "Faux" And estFacturable = "Vrai" Then
            total_hres_facturable = total_hres_facturable + hres
            h(3) = hres
        End If
        
        h(4) = 0
        If estDetruit = "Faux" And estFacturable = "Faux" Then
            total_hres_non_facturable = total_hres_non_facturable + hres
            h(4) = hres
        End If
        
        h(5) = 0
        If estDetruit = "Faux" And estFacturee = "Vrai" Then
            total_hres_facturees = total_hres_facturees + hres
            h(5) = hres
        End If
        
        'TEC = Heures Facturrables - Heures facturées
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
            dict_TEC_ID.add TECID, 0
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le TEC_ID '" & TECID & "' est un doublon pour la ligne '" & i & "'")
            r = r + 1
            cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        If dict_prof.Exists(prof & "-" & profID) = False Then
            dict_prof.add prof & "-" & profID, 0
        End If
        
        'Summary by Date
        d = day(dateTEC)
        m = month(dateTEC)
        y = year(dateTEC)
        keyDate = Format$(y, "0000") & Format$(m, "00") & Format$(d, "00") & Fn_Pad_A_String(prof, " ", 4, "L")
        p = InStr(bigStrDateProf, keyDate)
        If p = 0 Then
            rArr = rArr + 1
            pArr = rArr
            bigStrDateProf = bigStrDateProf & keyDate & Format$(rArr, "0000") & "|"
        Else
            pArr = Mid(bigStrDateProf, p + 12, 4)
        End If
        arrHres(pArr, 1) = arrHres(pArr, 1) + h(1)
        arrHres(pArr, 2) = arrHres(pArr, 2) + h(2)
        arrHres(pArr, 3) = arrHres(pArr, 3) + h(3)
        arrHres(pArr, 4) = arrHres(pArr, 4) + h(4)
        arrHres(pArr, 5) = arrHres(pArr, 5) + h(5)
        arrHres(pArr, 6) = arrHres(pArr, 6) + h(6)
    Next i
    
    Call SortDelimitedString(bigStrDateProf, "|")
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " charges de temps ont été analysées!")
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
    If cas_doublon_TECID = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de TEC_ID")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_TECID & " cas de doublons pour les TEC_ID")
        r = r + 1
    End If
    
    If cas_date_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune date INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        r = r + 1
    End If
    
    If cas_date_future = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune date dans le futur")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_date_future & " cas de date FUTURE")
        r = r + 1
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    r = r + 1
    
    If cas_hres_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune heures INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturable_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturable' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        r = r + 1
    End If
    
    If cas_estFacturee_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturee' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        r = r + 1
    End If
    
    If cas_estDetruit_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estDetruit' n'est INVALIDE")
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        r = r + 1
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La somme des heures donne ce résultat:")
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures inscrites       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures détruites       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures restantes       : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures facturables     : " & formattedHours)
    r = r + 1
    
    formattedHours = Format$(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures non_facturables : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures Facturées       : " & formattedHours)
    r = r + 1

    formattedHours = Format$(total_hres_facturable - total_hres_facturees, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures TEC             : " & formattedHours)
    r = r + 1

    Dim r2 As Integer
    r2 = 2 'Output to wsSommaire
    
    Dim components() As String
    components = Split(bigStrDateProf, "|")
    
    Dim dateStr As String
    For i = LBound(components) To UBound(components)
        dateStr = Left(components(i), 8)
        dateStr = DateSerial(Mid(dateStr, 1, 4), Mid(dateStr, 5, 2), Mid(dateStr, 7, 2))
        prof = Trim(Mid(components(i), 9, 4))
        pArr = CInt(Mid(components(i), 13, 4))
        wsSommaire.Cells(r2, 1).value = Format$(dateStr, "mm/dd/yyyy")
        wsSommaire.Cells(r2, 2).value = prof
        wsSommaire.Cells(r2, 3).value = arrHres(pArr, 1)
        wsSommaire.Cells(r2, 4).value = arrHres(pArr, 2)
        wsSommaire.Cells(r2, 5).value = arrHres(pArr, 3)
        wsSommaire.Cells(r2, 6).value = arrHres(pArr, 4)
        wsSommaire.Cells(r2, 7).value = arrHres(pArr, 5)
        wsSommaire.Cells(r2, 8).value = arrHres(pArr, 6)
        r2 = r2 + 1
    Next i
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set dict_TEC_ID = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modAppli:check_TEC", startTime)

End Sub

Sub ADMIN_DataFiles_Folder_Selection() '2024-03-28 @ 14:10

    Dim SharedFolder As FileDialog: Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With SharedFolder
        .Title = "Choisir le répertoire de données partagées, selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F5").value = .selectedItems(1)
        End If
    End With
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set SharedFolder = Nothing
    
End Sub

Sub ADMIN_Invoices_Excel_Folder_Selection() '2024-08-04 @ 07:30

    Dim SharedFolder As FileDialog: Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With SharedFolder
        .Title = "Choisir le répertoire des factures (Format Excel)"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F7").value = .selectedItems(1)
        End If
    End With
    
    'Cleaning memory - 2024-08-04 @ 07:28
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
            .size = 10
            .Italic = True
            .Bold = True
        End With
        .HorizontalAlignment = xlCenter
    End With
    
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
        .Title = "Choisir le répertoire des copies de facture (PDF), selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F6").value = .selectedItems(1)
        End If
    End With
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set PDFFolder = Nothing

End Sub

Sub Apply_Conditional_Formatting_Alternate(rng As Range, headerRows As Long, Optional EmptyLine As Boolean = False)

    Dim ws As Worksheet: Set ws = rng.Worksheet
    Dim DataRange As Range
    
    'Remove the worksheet conditional formatting
    ws.Cells.FormatConditions.delete
    
    'Determine the range excluding header rows
    Set DataRange = ws.Range(rng.Cells(headerRows + 1, 1), ws.Cells(ws.Cells(ws.rows.count, rng.Column).End(xlUp).Row, rng.columns.count))

    'Add the standard conditional formatting
    Dim formula As String
    If EmptyLine = False Then
        formula = "=ET($A2<>"""";MOD(LIGNE();2)=1)"
    Else
        formula = "=MOD(LIGNE();2)=1"
    End If
    
    DataRange.FormatConditions.add Type:=xlExpression, Formula1:= _
        formula
    DataRange.FormatConditions(DataRange.FormatConditions.count).SetFirstPriority
    With DataRange.FormatConditions(1).Font
        .Strikethrough = False
        .TintAndShade = 0
    End With
    With DataRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
    End With
    DataRange.FormatConditions(1).StopIfTrue = False

End Sub

Sub Apply_Worksheet_Format(ws As Worksheet, rng As Range, headerRow As Long)

    'Common stuff to all worksheets
    rng.EntireColumn.AutoFit 'Autofit all columns
    
    'Conditional Formatting (many steps)
    '1) Remove existing conditional formatting
        rng.Cells.FormatConditions.delete 'Remove the worksheet conditional formatting
    
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
            usedRange.FormatConditions.add Type:=xlExpression, _
                Formula1:="=ET($A2<>"""";mod(LIGNE();2)=1)"
    '        usedRange.FormatConditions.add Type:=xlExpression, _
    '            Formula1:="=ET($A2<>"""";MOD(LIGNE();2)=1)"
            usedRange.FormatConditions(usedRange.FormatConditions.count).SetFirstPriority
            With usedRange.FormatConditions(1).Font
                .Strikethrough = False
                .TintAndShade = 0
            End With
            With usedRange.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
            End With
            usedRange.FormatConditions(1).StopIfTrue = False
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
                .Range("A" & firstDataRow & ":M" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B" & firstDataRow & ":B" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
                .Range("C" & firstDataRow & ":C" & lastUsedRow & _
                     ", D" & firstDataRow & ":D" & lastUsedRow & _
                     ", E" & firstDataRow & ":E" & lastUsedRow & _
                     ", G" & firstDataRow & ":G" & lastUsedRow).HorizontalAlignment = xlLeft
                With .Range("I" & firstDataRow & ":N" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
                End With
            End With
       
        Case "wshDEB_Trans"
            With wshDEB_Trans
                .Range("A2:Q" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2:B" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
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
        
        Case "wshENC_Détails"
            With wshENC_Détails
                .Range("A2:A" & lastUsedRow & ", B2:B" & lastUsedRow & ", D2:D" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("C2:C" & lastUsedRow & ", E2:EB" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("E2:E" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("E2:E" & lastUsedRow).NumberFormat = "#,##0.00"
            End With
        
        Case "wshENC_Entête"
            With wshENC_Entête
                .Range("A2:A" & lastUsedRow & ", B2:B" & lastUsedRow & ", D2:D" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("C2:C" & lastUsedRow & ", E2:E" & lastUsedRow & ", G2:G" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("F2:F" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("F2:F" & lastUsedRow).NumberFormat = "#,##0.00$"
            End With
        
        Case "wshFAC_Comptes_Clients"
            With wshFAC_Comptes_Clients
                .Range("A" & firstDataRow & ":B" & lastUsedRow & ", " & _
                       "D" & firstDataRow & ":G" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("C" & firstDataRow & ":C" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("H" & firstDataRow & ":J" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("B" & firstDataRow & ":B" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
                .Range("G" & firstDataRow & ":G" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
                .Range("H" & firstDataRow & ":J" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("A1").CurrentRegion.EntireColumn.AutoFit
            End With
        
        Case "wshFAC_Détails"
            With usedRange
                .Range("A" & firstDataRow & ":A" & lastUsedRow & ", C" & firstDataRow & ":C" & lastUsedRow & ", F" & firstDataRow & ":F" & lastUsedRow & ", G" & firstDataRow & ":G" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B" & firstDataRow & ":B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("D" & firstDataRow & ":E" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("C" & firstDataRow & ":C" & lastUsedRow).NumberFormat = "#,##0.00"
                .Range("D" & firstDataRow & ":E" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("H" & firstDataRow & ":H" & lastUsedRow & ",J" & firstDataRow & ":J" & lastUsedRow & ",L" & firstDataRow & ":L" & lastUsedRow & ",N" & firstDataRow & ":T" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("O" & firstDataRow & ":O" & lastUsedRow & ",Q" & firstDataRow & ":Q" & lastUsedRow).NumberFormat = "#0.000 %"
            End With
        
        Case "wshFAC_Entête"
            With wshFAC_Entête
                .Range("A" & firstDataRow & ":D" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B" & firstDataRow & ":B" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
                .Range("E" & firstDataRow & ":I" & lastUsedRow & ",K" & firstDataRow & ":K" & lastUsedRow & ",M" & firstDataRow & ":M" & lastUsedRow & ",O" & firstDataRow & ":O" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("J" & firstDataRow & ":J" & lastUsedRow & ",L" & firstDataRow & ":L" & lastUsedRow & ",N" & firstDataRow & ":N" & lastUsedRow & ",P" & firstDataRow & ":V" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("J" & firstDataRow & ":J" & lastUsedRow & ",L" & firstDataRow & ":L" & lastUsedRow & ",N" & firstDataRow & ":N" & lastUsedRow & ",P" & firstDataRow & ":V" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("Q" & firstDataRow & ":Q" & lastUsedRow & ",S" & firstDataRow & ":S" & lastUsedRow).NumberFormat = "#0.000 %"
            End With

        Case "wshFAC_Projets_Détails"
            With wshFAC_Projets_Détails
                .Range("A" & firstDataRow & ":A" & lastUsedRow & ", C" & firstDataRow & ":G" & lastUsedRow & ", I" & firstDataRow & ":J" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B" & firstDataRow & ":B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("H" & firstDataRow & ":I" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("H" & firstDataRow & ":H" & lastUsedRow).NumberFormat = "#,##0.00"
            End With
        
        Case "wshFAC_Projets_Entête"
            With wshFAC_Projets_Entête
                .Range("A" & firstDataRow & ":A" & lastUsedRow & ", C" & firstDataRow & ":D" & lastUsedRow & ", F" & firstDataRow & ":F" & lastUsedRow & _
                       ", J" & firstDataRow & ":J" & lastUsedRow & ", N" & firstDataRow & ":N" & lastUsedRow & ", R" & firstDataRow & ":R" & lastUsedRow & _
                       ", V" & firstDataRow & ":V" & lastUsedRow & ", Z" & firstDataRow & ":AA" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B" & firstDataRow & ":B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("E" & firstDataRow & ":E" & lastUsedRow & ", I" & firstDataRow & ":I" & lastUsedRow & ", M" & firstDataRow & ":M" & lastUsedRow & _
                        ", Q" & firstDataRow & ":Q" & lastUsedRow & ", U" & firstDataRow & ":U" & lastUsedRow & ", Y" & firstDataRow & ":Y" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("G" & firstDataRow & ":H" & lastUsedRow).NumberFormat = "#,##0.00"
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
                .Range("A" & firstDataRow & ":J" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B" & firstDataRow & ":B" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
                .Range("C" & firstDataRow & ":C" & lastUsedRow & _
                    ", D" & firstDataRow & ":D" & lastUsedRow & _
                    ", F" & firstDataRow & ":F" & lastUsedRow & _
                    ", I" & firstDataRow & ":I" & lastUsedRow) _
                        .HorizontalAlignment = xlLeft
                With .Range("G" & firstDataRow & ":H" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
                End With
                With .Range("A" & firstDataRow & ":A" & lastUsedRow) _
                    .Range("J" & firstDataRow & ":J" & lastUsedRow).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent5
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            End With
        
        Case "wshTEC_Local"
            With wshTEC_Local
                .Range("A" & firstDataRow & ":P" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("F" & firstDataRow & ":F" & lastUsedRow & ", G" & firstDataRow & _
                                            ":G" & lastUsedRow & ", I" & firstDataRow & _
                                            ":I" & lastUsedRow & ", O" & firstDataRow & _
                                            ":O" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("H" & firstDataRow & ":H" & lastUsedRow).NumberFormat = "#0.00"
                .Range("K" & firstDataRow & ":K" & lastUsedRow).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                .columns("F").ColumnWidth = 45
                .columns("G").ColumnWidth = 65
                .columns("I").ColumnWidth = 25
            End With

    End Select

End Sub

Sub Compare_2_Workbooks_Column_Formatting()                      '2024-08-19 @ 16:24

    'Erase and create a new worksheet for differences
    Dim wsDiff As Worksheet
    Call CreateOrReplaceWorksheet("Différences_Colonnes")
    Set wsDiff = ThisWorkbook.Worksheets("Différences_Colonnes")
    wsDiff.Range("A1").value = "Worksheet"
    wsDiff.Range("B1").value = "Nb. colonnes"
    wsDiff.Range("C1").value = "Colonne"
    wsDiff.Range("D1").value = "Valeur originale"
    wsDiff.Range("E1").value = "Nouvelle valeur"
    Call Make_It_As_Header(wsDiff.Range("A1:E1"))

    'Set your workbooks and worksheets here
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\GCF_DataFiles\GCF_BD_MASTER_COPY.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx")
    
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
                                        " colonnes analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:E" & diffRow)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:E" & diffRow)
    Dim header1 As String: header1 = wb1.name & " vs. " & wb2.name
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsDiff, rngToPrint, header1, header2, "P")
    
    'Close the 2 workbooks without saving anything
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    'Output differences
    If diffLog <> "" Then
        MsgBox "Différences trouvées:" & vbCrLf & diffLog
    Else
        MsgBox "Aucune différence dans les colonnes."
    End If
    
    'Cleanup
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
    Call CreateOrReplaceWorksheet("Différences_Lignes")
    Set wsDiff = ThisWorkbook.Worksheets("Différences_Lignes")
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
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\GCF_DataFiles\GCF_BD_MASTER_COPY.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx")
    
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
            wsDiff.Cells(diffRow, 6).value = "Le client a ajouté " & nbRowProd - nbRowDev & " lignes dans la feuille"
            nbRow = nbRowProd
        End If
        If nbRowProd < nbRowDev Then
            wsDiff.Cells(diffRow, 6).value = "Le dev a ajouté " & nbRowDev - nbRowProd & " lignes dans la feuille"
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
                                        " lignes analysées dans l'ensemble du Workbook ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:I" & diffRow)
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True)

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:I" & diffRow)
    Dim header1 As String: header1 = wb1.name & " vs. " & wb2.name
    Dim header2 As String: header2 = "Changements de lignes ou cellules"
    Call Simple_Print_Setup(wsDiff, rngToPrint, header1, header2, "P")
    
    'Close the 2 workbooks without saving anything
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    'Output differences
    If diffLogMess <> "" Then
        MsgBox "Différences trouvées:" & vbCrLf & diffLogMess
    Else
        MsgBox "Aucune différence dans les lignes."
    End If
    
    'Cleanup
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
