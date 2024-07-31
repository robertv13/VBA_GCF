Attribute VB_Name = "modAppli_Utils"
Option Explicit

Sub Clone_Last_Line_Formatting_For_New_Records(workbookPath As String, wSheet As String, numberRows As Long)

    'Open the workbook
    Dim wb As Workbook: Set wb = Workbooks.Open(workbookPath)
    Dim ws As Worksheet: Set ws = wb.Sheets(wSheet)

    'Find the last row with data in column A
    Dim lastRow As Long
    lastRow = ws.Range("A9999").End(xlUp).row
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

'Sub Printer_Page_Setup(ws As Worksheet, _
'                       rng As Range, _
'                       header1 As String, _
'                       header2 As String, _
'                       Optional Orient As String = "L") '2024-07-14 @ 06:51
'
'    Dim retries As Integer
'
'    On Error GoTo CleanUp
'
'    'Retry loop for setting PrintCommunication to False
'    retries = 3
'    Do While retries > 0
'        On Error Resume Next
'        Application.PrintCommunication = False
'        If Err.Number = 0 Then Exit Do
'        retries = retries - 1
'        Application.Wait (Now + TimeValue("0:00:01"))
'        On Error GoTo CleanUp
'    Loop
'
'    If retries = 0 Then
'        MsgBox "Failed to set PrintCommunication to False after multiple attempts", vbCritical
'        Exit Sub
'    End If
'
'    With ws.PageSetup
'        .PrintArea = rng.Address
'        .PrintTitleRows = "$1:$1"
'        .PrintTitleColumns = ""
'
'        .LeftHeader = ""
'        .CenterHeader = "&""-,Gras""&14&K0070C0" & header1 & Chr(10) & header2
'        .RightHeader = ""
'        .LeftFooter = "&11&D - &T"
'        .CenterFooter = "&11&KFF0000&A"
'        .RightFooter = "&11Page &P de &N"
'
'        .TopMargin = Application.InchesToPoints(0.55)
'        .LeftMargin = Application.InchesToPoints(0.15)
'        .RightMargin = Application.InchesToPoints(0.15)
'        .HeaderMargin = Application.InchesToPoints(0.15)
'        .FooterMargin = Application.InchesToPoints(0.15)
'        .BottomMargin = Application.InchesToPoints(0.4)
'
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .PrintQuality = 600
'        .CenterHorizontally = True
'        .CenterVertically = False
'        If Orient = "L" Then
'            .Orientation = xlLandscape
'        Else
'            .Orientation = xlPortrait
'        End If
'        .Draft = False
'        .PaperSize = xlPaperLetter
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = 100
'        .FitToPagesWide = 1
'        .FitToPagesTall = False
'        .PrintErrors = xlPrintErrorsDisplayed
'        .OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = True
'        .AlignMarginsHeaderFooter = True
'
'        'Clear EvenPage headers and footers if they exist
'        On Error Resume Next
'       .EvenPage.LeftHeader.text = ""
'        .EvenPage.CenterHeader.text = ""
'        .EvenPage.RightHeader.text = ""
'        .EvenPage.LeftFooter.text = ""
'        .EvenPage.CenterFooter.text = ""
'        .EvenPage.RightFooter.text = ""
'
'        .FirstPage.LeftHeader.text = ""
'        .FirstPage.CenterHeader.text = ""
'        .FirstPage.RightHeader.text = ""
'        .FirstPage.LeftFooter.text = ""
'        .FirstPage.CenterFooter.text = ""
'        .FirstPage.RightFooter.text = ""
'    End With
'
'CleanUp:
'    'Retry loop for setting PrintCommunication to True
'    retries = 3
'    Do While retries > 0
'        On Error Resume Next
'        Application.PrintCommunication = True
'        If Err.Number = 0 Then Exit Do
'        retries = retries - 1
'        Application.Wait (Now + TimeValue("0:00:01"))
'        On Error GoTo CleanUp
'    Loop
'    If retries = 0 Then
'        MsgBox "Failed to set PrintCommunication to True after multiple attempts", vbCritical
'    End If
'    On Error GoTo 0
'
'End Sub

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
    rng.Parent.Protect UserInterfaceOnly:=True


End Sub

Public Sub UnprotectCells(rng As Range)

    'Lock the checkbox
    rng.Locked = False
    
    'Protect the worksheet
    rng.Parent.Protect UserInterfaceOnly:=True


End Sub

Sub Start_Routine(subName As String) '2024-06-06 @ 10:12

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
            lastUsedRow = .Range("A99999").End(xlUp).row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("B" & lastUsedRow).value = subName & " - entering"
        End With
    End If

End Sub

Sub Output_Timer_Results(subName As String, t As Double)

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
            lastUsedRow = .Range("A99999").End(xlUp).row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("B" & lastUsedRow).value = subName
            If t Then
                .Range("C" & lastUsedRow).value = Round(Timer - t, 4)
                .Range("C" & lastUsedRow).NumberFormat = "#,##0.0000"
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
    Set ws = ThisWorkbook.Worksheets.add
    ws.name = wsName

    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
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

    Application.ScreenUpdating = False
    
    Call Erase_And_Create_Worksheet("Analyse_Intégrité")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    wsOutput.Range("A1").value = "Feuille"
    wsOutput.Range("B1").value = "Message"
    wsOutput.Range("C1").value = "TimeStamp"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"))

    'Data starts at row 2
    Dim r As Long: r = 2
    Dim readRows As Long
    
    'wshBD_Clients --------------------------------------------------------------- Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "BD_Clients")
    r = r + 1
    
    Call Client_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_Clients(r, readRows)

    'wshBD_Fournisseurs ----------------------------------------------------- Fournisseurs
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Fournisseurs")
    r = r + 1
    
    Call Fournisseur_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_Fournisseurs(r, readRows)
    
    'wshFAC_Détails ---------------------------------------------------------- FAC_Détails
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Détails")
    r = r + 1
    
    Call FAC_Détails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Détails a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Détails(r, readRows)
    
    'wshFAC_Entête ------------------------------------------------------------ FAC_Entête
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Entête")
    r = r + 1
    
    Call FAC_Entête_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Entête a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Entête(r, readRows)
    
    'wshFAC_Projets_Détails ------------------------------------------ FAC_Projets_Détails
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Projets_Détails")
    r = r + 1
    
    Call FAC_Projets_Détails_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Projets_Détails a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Projets_Détails(r, readRows)
    
    'wshFAC_Projets_Entête -------------------------------------------- FAC_Projets_Entête
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "FAC_Projets_Entête")
    r = r + 1
    
    Call FAC_Projets_Entête_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "FAC_Projets_Entête a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_FAC_Projets_Entête(r, readRows)
    
    'wshGL_Trans ---------------------------------------------------------------- GL_Trans
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "GL_Trans")
    r = r + 1
    
    Call GL_Trans_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "GL_Trans a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1

    Call check_GL_Trans(r, readRows)
    
    'wshTEC_DB_Data ---------------------------------------------------------- TEC_DB_Data
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "TEC_TDB_Data")
    r = r + 1
    
    Call check_TEC_TDB_Data(r, readRows)
    
    'wshTEC_Local -------------------------------------------------------------- TEC_Local
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "TEC_Local")
    r = r + 1
    
    Call TEC_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "TEC_Local a été importée du fichier BD_MASTER.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_TEC(r, readRows)
    
    'Adjust the Output Worksheet
    With wsOutput.Range("A2:C" & r).Font
        .name = "Courier New"
        .Size = 10
    End With
    
    wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit

   'Result print setup - 2024-07-20 @ 14:31
    Dim lastUsedRow As Long
    lastUsedRow = r + 1
    wsOutput.Range("A" & lastUsedRow).value = "**** " & Format$(readRows, "###,##0") & _
                                    " lignes analysées dans l'ensemble de l'application ***"
    
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "Vérification d'intégrité des tables"
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "P")
    
    ThisWorkbook.Worksheets("Analyse_Intégrité").Activate
    
    MsgBox "La vérification d'intégrité est terminé" & vbNewLine & vbNewLine & "Voir la feuille 'Analyse_Intégrité'", vbInformation
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_Clients(ByRef r As Long, ByRef readRows As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshBD_Clients
    Dim ws As Worksheet: Set ws = wshBD_Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(ws.usedRange.rows.count - 1, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshBD_Clients'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        If dict_code_client.Exists(code) = False Then
            dict_code_client.add code, nom
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le code '" & code & "' est un doublon pour le client '" & nom & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
    Next i
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 1, "##,##0") & " clients ont été analysés!")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de nom")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de code")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    r = r + 1
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_Fournisseurs(ByRef r As Long, ByRef readRows As Long)
    
    Application.ScreenUpdating = False

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshBD_fournisseurs
    Dim ws As Worksheet: Set ws = wshBD_Fournisseurs
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(ws.usedRange.rows.count - 1, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshBD_Fournisseurs'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        If dict_code_fournisseur.Exists(code) = False Then
            dict_code_fournisseur.add code, nom
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le code '" & code & "' est un doublon pour le nom '" & nom & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
    Next i
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 1, "#,##0") & " fournisseurs ont été analysés!")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1)
    
    If cas_doublon_nom = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de nom")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de code")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    r = r + 1
    
Clean_Exit:
    'Cleaning memory - 2024-07-04 @ 12:37
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_FAC_Détails(ByRef r As Long, ByRef readRows As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshFAC_Détails
    Dim ws As Worksheet: Set ws = wshFAC_Détails
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Entête
    
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A" & 1 + headerRow & ":A" & lastUsedRow)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Détails'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    'Transfer data from Worksheet into an Array (arr)
    Dim arr As Variant
    arr = wshFAC_Détails.Range("A1").CurrentRegion.Offset(1, 0).value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String, oldInv_No As String
    Dim result As Variant
    For i = LBound(arr, 1) + 2 To UBound(arr, 1) - 1 'Two lines of header !
        Inv_No = arr(i, 1)
        If Inv_No <> oldInv_No Then
            result = Application.WorksheetFunction.XLookup(ws.Cells(i, 1).value, _
                                                       rngMaster, _
                                                       rngMaster, _
                                                       "Not Found", _
                                                       0, _
                                                       1)

            result = Application.WorksheetFunction.XLookup(ws.Cells(i, 1), rngMaster, rngMaster, "Not Found", 0, 1)
            oldInv_No = Inv_No
        End If
        If result = "Not Found" Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " n'existe pas dans FAC_Entête")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 3)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " le nombre d'heures est INVALIDE '" & arr(i, 3) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " le taux horaire est INVALIDE '" & arr(i, 5) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " le montant est INVALIDE '" & arr(i, 5) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - 2, "##,##0") & " lignes de transactions ont été analysées")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
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
    
End Sub

Private Sub check_FAC_Entête(ByRef r As Long, ByRef readRows As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Entête'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim arr As Variant
    arr = wshFAC_Entête.Range("A1").CurrentRegion.Offset(2, 0).Resize(lastUsedRow - headerRow, ws.Range("A1").CurrentRegion.columns.count - headerRow).value
    If UBound(arr, 1) < 3 Then
        r = r + 1
        GoTo Clean_Exit
    End If

    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim Inv_No As String
    For i = LBound(arr, 1) + 2 To UBound(arr, 1) 'Two lines of header !
        Inv_No = arr(i, 1)
        If IsDate(arr(i, 2)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La facture '" & Inv_No & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 2) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1), "##,##0") & " lignes de transactions ont été analysées")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_FAC_Projets_Détails(ByRef r As Long, ByRef readRows As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshFAC_Projets_Détails
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 2
        GoTo Clean_Exit
    End If

    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim wsMaster As Worksheet: Set wsMaster = wshFAC_Projets_Entête
    lastUsedRow = wsMaster.Range("A99999").End(xlUp).row
    Dim rngMaster As Range: Set rngMaster = wsMaster.Range("A2:A" & lastUsedRow)
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Projets_Détails'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    'Transfer data from Worksheet into an Array (arr)
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.rows.count - 1
    If numRows < 1 Then
        r = r + 1
        GoTo Clean_Exit
    End If
    
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Offset(1, 0).Resize(numRows, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim projetID As String, oldProjetID As String
    Dim lookUpValue As String, result As Variant
    For i = LBound(arr, 1) To UBound(arr, 1) - 1 'One line of header !
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 3)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le ClientID est INVALIDE '" & arr(i, 3) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le TECID est INVALIDE '" & arr(i, 4) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le ProfID est INVALIDE '" & arr(i, 5) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Heures sont INVALIDES '" & arr(i, 8) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " lignes ont été analysées")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
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
    
End Sub

Private Sub check_FAC_Projets_Entête(ByRef r As Long, ByRef readRows As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshFAC_Projets_Entête'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    'Establish the number of rows before transferring it to an Array
    Dim numRows As Long
    numRows = ws.Range("A1").CurrentRegion.rows.count
    If numRows <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 2
        GoTo Clean_Exit
    End If
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Offset(1, 0).Resize(numRows, ws.Range("A1").CurrentRegion.columns.count).value
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long
    Dim projetID As String
    For i = LBound(arr, 1) To UBound(arr, 1) 'One line of header !
        projetID = arr(i, 1)
        If IsNumeric(arr(i, 3)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le ClientID est INVALIDE '" & arr(i, 3) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsDate(arr(i, 4)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " la date est INVALIDE '" & arr(i, 4) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 5)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le total des honoraires est INVALIDE '" & arr(i, 5) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 7)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du premier sommaire sont INVALIDES '" & arr(i, 7) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 8)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du premier sommaire est INVALIDE '" & arr(i, 8) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If IsNumeric(arr(i, 9)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du premier sommaire sont INVALIDES '" & arr(i, 9) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 11) <> "" And IsNumeric(arr(i, 11)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du second sommaire sont INVALIDES '" & arr(i, 11) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 12) <> "" And IsNumeric(arr(i, 12)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du second sommaire est INVALIDE '" & arr(i, 12) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 13) <> "" And IsNumeric(arr(i, 13)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du second sommaire sont INVALIDES '" & arr(i, 13) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 15) <> "" And IsNumeric(arr(i, 15)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du troisième sommaire sont INVALIDES '" & arr(i, 15) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 16) <> "" And IsNumeric(arr(i, 16)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du troisième sommaire est INVALIDE '" & arr(i, 16) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 17) <> "" And IsNumeric(arr(i, 17)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du troisième sommaire sont INVALIDES '" & arr(i, 17) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 19) <> "" And IsNumeric(arr(i, 19)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du quatrième sommaire sont INVALIDES '" & arr(i, 19) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 20) <> "" And IsNumeric(arr(i, 20)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du quatrième sommaire est INVALIDE '" & arr(i, 20) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 21) <> "" And IsNumeric(arr(i, 21)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du quatrième sommaire sont INVALIDES '" & arr(i, 21) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 23) <> "" And IsNumeric(arr(i, 23)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les heures du cinquième sommaire sont INVALIDES '" & arr(i, 23) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 24) <> "" And IsNumeric(arr(i, 24)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " le taux horaire du cinquième sommaire est INVALIDE '" & arr(i, 24) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        If arr(i, 25) <> "" And IsNumeric(arr(i, 25)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le projet '" & projetID & "' à la ligne " & i & " les Honoraires du cinquième sommaire sont INVALIDES '" & arr(i, 25) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " lignes de transactions ont été analysées")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 2
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_GL_Trans(ByRef r As Long, ByRef readRows As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshGL_Trans
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(ws.usedRange.columns.count, "#,##0") & " colonnes dans cette table")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshGL_Trans'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    On Error Resume Next
    Dim planComptable As Range: Set planComptable = wshAdmin.Range("dnrPlanComptable_All")
    On Error GoTo 0

    If planComptable Is Nothing Then
        MsgBox "La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée ou est INVALIDE!", vbExclamation
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La plage nommée 'dnrPlanComptable_All' n'a pas été trouvée!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
        Exit Sub
    End If
    
    Dim strCodeGL As String, strDescGL As String
    Dim ligne As Range
    For Each ligne In planComptable.rows
        strCodeGL = strCodeGL + ligne.Cells(1, 2).value + "|:|"
        strDescGL = strDescGL + ligne.Cells(1, 1).value + "|:|"
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
    Dim dt As Double, ct As Double
    Dim GL_Entry_No As String, glCode As String, glDescr As String
    Dim result As Variant
    For i = LBound(arr, 1) To UBound(arr, 1)
        GL_Entry_No = arr(i, 1)
        If dict_GL_Entry.Exists(GL_Entry_No) = False Then
            dict_GL_Entry.add GL_Entry_No, row
            sum_arr(row, 1) = GL_Entry_No
            row = row + 1
        End If
        If IsDate(arr(i, 2)) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** L'écriture #  " & GL_Entry_No & " ' à la ligne " & i & " a une date INVALIDE '" & arr(i, 2) & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        glCode = arr(i, 5)
        If InStr(1, strCodeGL, glCode + "|:|") = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le compte '" & glCode & "' à la ligne " & i & " est INVALIDE '")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        glDescr = arr(i, 6)
        If InStr(1, strDescGL, glDescr + "|:|") = 0 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** La description du compte '" & glDescr & "' à la ligne " & i & " est INVALIDE")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        dt = arr(i, 7)
        If IsNumeric(dt) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le montant du débit '" & dt & "' à la ligne " & i & " n'est pas une valeur numérique")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        ct = arr(i, 8)
        If IsNumeric(ct) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le montant du débit '" & ct & "' à la ligne " & i & " n'est pas une valeur numérique")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
        End If
        currentRow = dict_GL_Entry(GL_Entry_No)
        sum_arr(currentRow, 2) = sum_arr(currentRow, 2) + dt
        sum_arr(currentRow, 3) = sum_arr(currentRow, 3) + ct
        If arr(i, 10) <> "" Then
            If IsDate(arr(i, 10)) = False Then
                Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le TimeStamp '" & arr(i, 10) & "' à la ligne " & i & " n'est pas une date VALIDE")
                Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_hors_balance = cas_hors_balance + 1
        End If
        sum_dt = sum_dt + dt
        sum_ct = sum_ct + ct
    Next v
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " lignes de transactions ont été analysées")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Un total de " & dict_GL_Entry.count & " écritures ont été analysées")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    If cas_hors_balance = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Chacune des écritures balancent au niveau de l'écriture")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_hors_balance & " écriture(s) qui ne balance(nt) pas !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Les totaux des transactions sont:")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Dt = " & Format$(sum_dt, "###,###,##0.00 $"))
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Ct = " & Format$(sum_ct, "###,###,##0.00 $"))
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    If sum_dt - sum_ct <> 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Hors-Balance de " & Format$(sum_dt - sum_ct, "###,###,##0.00$"))
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    r = r + 1
    
Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set planComptable = Nothing
    Set v = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_TEC_TDB_Data(ByRef r As Long, ByRef readRows As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshTEC_DB_Data
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    Dim headerRow As Long: headerRow = 1
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim lastUsedCol As Long
    lastUsedCol = ws.Range("A2").End(xlToRight).Column
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshTEC_TDB_Data'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_date_invalide = cas_date_invalide + 1
        Else
            If dateTEC < minDate Then minDate = dateTEC
            If dateTEC > maxDate Then maxDate = dateTEC
        End If
        hres = arr(i, 5)
        If IsNumeric(hres) = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_hres_invalide = cas_hres_invalide + 1
        End If
        estFacturable = arr(i, 6)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If
        estFacturee = arr(i, 7)
        If InStr("Vrai^Faux^", estFacturee & "^") = 0 Or Len(estFacturee) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        estDetruit = arr(i, 8)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        If dict_prof.Exists(prof & "-" & profID) = False Then
            dict_prof.add prof & "-" & profID, 0
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " charges de temps ont été analysées!")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
    If cas_doublon_TECID = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de TEC_ID")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_TECID & " cas de doublons pour les TEC_ID")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    
    If cas_date_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune date INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    r = r + 1
    
    If cas_hres_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune heures INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    
    If cas_estFacturable_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturable' n'est INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    
    If cas_estFacturee_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturee' n'est INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    
    If cas_estDetruit_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estDetruit' n'est INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La somme des heures donne ce resultat:")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures inscrites       : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures détruites       : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format$(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures restantes       : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format$(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures facturables     : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format$(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures non_facturables : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 2

Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_TEC(ByRef r As Long, ByRef readRows As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshTEC_Local
    Dim ws As Worksheet: Set ws = wshTEC_Local
    Dim headerRow As Long: headerRow = 2
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow <= headerRow Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Cette feuille est vide !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 2
        GoTo Clean_Exit
    End If
    
    Dim lastUsedCol As Long
    lastUsedCol = ws.Range("A2").End(xlToRight).Column
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Il y a " & Format$(lastUsedRow - headerRow, "###,##0") & _
        " lignes et " & Format$(lastUsedCol, "#,##0") & " colonnes dans cette table")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshTEC_Local'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion.Offset(2)
    Dim dict_TEC_ID As New Dictionary
    Dim dict_prof As New Dictionary
    
    Dim i As Long, TECID As Long, profID As String, prof As String, dateTEC As Date, testDate As Boolean
    Dim minDate As Date, maxDate As Date
    Dim code As String, nom As String, hres As Double, testHres As Boolean, estFacturable As Boolean
    Dim estFacturee As Boolean, estDetruit As Boolean
    Dim cas_doublon_TECID As Long, cas_date_invalide As Long, cas_doublon_prof As Long, cas_doublon_client As Long
    Dim cas_hres_invalide As Long, cas_estFacturable_invalide As Long, cas_estFacturee_invalide As Long
    Dim cas_estDetruit_invalide As Long
    Dim total_hres_inscrites As Double, total_hres_detruites As Double, total_hres_facturees As Double
    Dim total_hres_facturable As Double, total_hres_TEC As Double, total_hres_non_facturable As Double
    
    minDate = "12/31/2999"
    For i = LBound(arr, 1) To UBound(arr, 1) - 2
        TECID = arr(i, 1)
        profID = arr(i, 2)
        prof = arr(i, 3)
        dateTEC = arr(i, 4)
        testDate = IsDate(dateTEC)
        If testDate = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "***** TEC_ID =" & TECID & " a une date INVALIDE '" & dateTEC & " !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_date_invalide = cas_date_invalide + 1
        Else
            If dateTEC < minDate Then minDate = dateTEC
            If dateTEC > maxDate Then maxDate = dateTEC
        End If
        code = arr(i, 5)
        nom = arr(i, 6)
        hres = arr(i, 8)
        testHres = IsNumeric(hres)
        If testHres = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_hres_invalide = cas_hres_invalide + 1
        End If
        estFacturable = arr(i, 10)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If
        estFacturee = arr(i, 12)
        If InStr("Vrai^Faux^", estFacturee & "^") = 0 Or Len(estFacturee) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        estDetruit = arr(i, 14)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
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
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        If dict_prof.Exists(prof & "-" & profID) = False Then
            dict_prof.add prof & "-" & profID, 0
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format$(UBound(arr, 1) - headerRow, "##,##0") & " charges de temps ont été analysées!")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    'Add number of rows processed (read)
    readRows = readRows + UBound(arr, 1) - headerRow
    
    If cas_doublon_TECID = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de TEC_ID")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_TECID & " cas de doublons pour les TEC_ID")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    
    If cas_date_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune date INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     La date MINIMALE est '" & Format$(minDate, "dd/mm/yyyy") & "'")
    r = r + 1
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     La date MAXIMALE est '" & Format$(maxDate, "dd/mm/yyyy") & "'")
    r = r + 1
    
    If cas_hres_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune heures INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    
    If cas_estFacturable_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturable' n'est INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    
    If cas_estFacturee_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturee' n'est INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    
    If cas_estDetruit_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estDetruit' n'est INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La somme des heures donne ce resultat:")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format$(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures inscrites       : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format$(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures détruites       : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format$(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures restantes       : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format$(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures facturables     : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format$(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures non_facturables : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format$(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1

Clean_Exit:
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
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
            .Size = 10
            .Italic = True
            .Bold = True
        End With
        .HorizontalAlignment = xlCenter
    End With
    
End Sub

Sub Add_Message_To_WorkSheet(ws As Worksheet, r As Long, c As Long, m As String)

    ws.Cells(r, c).value = m

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

Sub CopyToClipboard() '2024-07-06 @ 07:37

    'Reference:
    
    Dim objData As Object
    Dim strText As String
    
    ' Create a new DataObject
    Set objData = CreateObject("MSForms.DataObject")
    
    ' Text to copy
    strText = "Hello, Clipboard!"
    
    ' Set text to DataObject
    objData.SetText strText
    ' Put text to Clipboard
    objData.PutInClipboard
    
    MsgBox "Text copied to clipboard!"
End Sub

Sub PasteFromClipboard() '2024-07-06 @ 07:37

    Dim objData As Object
    Dim strText As String
    
    ' Create a new DataObject
    Set objData = CreateObject("MSForms.DataObject")
    
    ' Get data from Clipboard
    objData.GetFromClipboard
    ' Get text from DataObject
    strText = objData.GetText
    
    ' Paste text into cell A1
    Range("A1").value = strText
    
    MsgBox "Text pasted from clipboard!"

End Sub

Sub Apply_Conditional_Formatting_Alternate(rng As Range, headerRows As Long, Optional EmptyLine As Boolean = False)

    Dim ws As Worksheet: Set ws = rng.Worksheet
    Dim dataRange As Range
    
    'Remove the worksheet conditional formatting
    ws.Cells.FormatConditions.delete
    
    'Determine the range excluding header rows
    Set dataRange = ws.Range(rng.Cells(headerRows + 1, 1), ws.Cells(ws.Cells(ws.rows.count, rng.Column).End(xlUp).row, rng.columns.count))

    'Add the standard conditional formatting
    Dim formula As String
    If EmptyLine = False Then
        formula = "=ET($A2<>"""";MOD(LIGNE();2)=1)"
    Else
        formula = "=MOD(LIGNE();2)=1"
    End If
    
    dataRange.FormatConditions.add Type:=xlExpression, Formula1:= _
        formula
    dataRange.FormatConditions(dataRange.FormatConditions.count).SetFirstPriority
    With dataRange.FormatConditions(1).Font
        .Strikethrough = False
        .TintAndShade = 0
    End With
    With dataRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
    End With
    dataRange.FormatConditions(1).StopIfTrue = False

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
                .Range("A2" & firstDataRow & ":P" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B2" & firstDataRow & ":B" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
                .Range("C2" & firstDataRow & ":C" & lastUsedRow & _
                     ", D2" & firstDataRow & ":D" & lastUsedRow & _
                     ", F2" & firstDataRow & ":F" & lastUsedRow & _
                     ", H2" & firstDataRow & ":H" & lastUsedRow & _
                     ", O2" & firstDataRow & ":O" & lastUsedRow).HorizontalAlignment = xlLeft
                With .Range("J2" & firstDataRow & ":N" & lastUsedRow)
                    .HorizontalAlignment = xlRight
                    .NumberFormat = "#,##0.00 $"
                End With
                .Range("A1").CurrentRegion.EntireColumn.AutoFit
            End With
        
        Case "wshENC_Détails"
            With wshENC_Détails
                .Range("A4" & firstDataRow & ":A" & lastUsedRow & ", C4" & firstDataRow & ":C" & lastUsedRow & ", F4" & firstDataRow & ":F" & lastUsedRow & ", G4" & firstDataRow & ":G" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("B4" & firstDataRow & ":B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("D4" & firstDataRow & ":E" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("C4" & firstDataRow & ":C" & lastUsedRow).NumberFormat = "#,##0.00"
                .Range("D4" & firstDataRow & ":E" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("H4" & firstDataRow & ":H" & lastUsedRow & ",J4" & firstDataRow & ":J" & lastUsedRow & ",L4" & firstDataRow & ":L" & lastUsedRow & ",N4" & firstDataRow & ":T" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("O4" & firstDataRow & ":O" & lastUsedRow & ",Q4" & firstDataRow & ":Q" & lastUsedRow).NumberFormat = "#0.000 %"
            End With
        
        Case "wshENC_Entête"
            With wshENC_Entête
                .Range("A2" & firstDataRow & ":F" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("A2" & firstDataRow & ":B" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("E2" & firstDataRow & ":E" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("E2" & firstDataRow & ":E" & lastUsedRow).NumberFormat = "#,##0.00$"
            End With
        
        Case "wshFAC_Comptes_Clients"
            With wshFAC_Comptes_Clients
                .Range("A3" & firstDataRow & ":B" & lastUsedRow & ", D3" & firstDataRow & ":F" & lastUsedRow & ", J3" & firstDataRow & ":J" & lastUsedRow).HorizontalAlignment = xlCenter
                .Range("C3" & firstDataRow & ":C" & lastUsedRow).HorizontalAlignment = xlLeft
                .Range("G3" & firstDataRow & ":I" & lastUsedRow).HorizontalAlignment = xlRight
                .Range("B3" & firstDataRow & ":B" & lastUsedRow).NumberFormat = "dd/mm/yyyy"
                .Range("G3" & firstDataRow & ":I" & lastUsedRow).NumberFormat = "#,##0.00 $"
                .Range("A1").CurrentRegion.EntireColumn.AutoFit
            End With
        
        Case "wshGL_EJ_Recurrente"
            With wshGL_EJ_Recurrente
                Union(.Range("C" & firstDataRow & ":C" & lastUsedRow), _
                      .Range("E" & firstDataRow & ":E" & lastUsedRow)).HorizontalAlignment = xlCenter
                Union(.Range("D" & firstDataRow & ":D" & lastUsedRow), _
                      .Range("F" & firstDataRow & ":F" & lastUsedRow), _
                      .Range("I" & firstDataRow & ":I" & lastUsedRow)).HorizontalAlignment = xlLeft
                With .Range("G" & firstDataRow & ":H" & lastUsedRow)
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
            End With

    End Select

End Sub


