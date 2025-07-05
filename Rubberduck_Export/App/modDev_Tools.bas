Attribute VB_Name = "modDev_Tools"
'@IgnoreModule UnassignedVariableUsage

Option Explicit

Sub Get_Range_From_Dynamic_Named_Range(dynamicRangeName As String, ByRef rng As Range)
    
    On Error Resume Next
    'Récupérer la formule associée au nom
    Dim refersToFormula As String
    refersToFormula = ThisWorkbook.Names(dynamicRangeName).RefersTo
    On Error GoTo 0
    
    If refersToFormula = "" Then
        MsgBox "La plage nommée '" & dynamicRangeName & "' n'existe pas ou est invalide.", vbExclamation
        Exit Sub
    End If
    
    'Tester et évaluer la plage
    On Error Resume Next
    Set rng = Application.Evaluate(refersToFormula)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Impossible de résoudre la plage nommée dynamique '" & dynamicRangeName & "'. Vérifiez la définition.", vbExclamation
        Exit Sub
    End If
    
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
                    circRef = circRef & ws.Name & "!" & cell.Address & vbCrLf
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
    
    'Libérer la mémoire
    Set cell = Nothing
    Set formulaCells = Nothing
    Set ws = Nothing
    
End Sub

Sub Compare_2_Workbooks_Column_Formatting()                      '2024-08-19 @ 16:24

    'Erase and create a new worksheet for differences
    Dim wsDiff As Worksheet
    Call CreateOrReplaceWorksheet("Différences_Colonnes")
    Set wsDiff = ThisWorkbook.Worksheets("Différences_Colonnes")
    wsDiff.Range("A1").Value = "Worksheet"
    wsDiff.Range("B1").Value = "Nb. colonnes"
    wsDiff.Range("C1").Value = "Colonne"
    wsDiff.Range("D1").Value = "Valeur originale"
    wsDiff.Range("E1").Value = "Nouvelle valeur"
    Call Make_It_As_Header(wsDiff.Range("A1:E1"), RGB(0, 112, 192))

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
        wsName = wso.Name
        Set wsn = wb2.Sheets(wsName)
        
        Dim nbCol As Integer
        nbCol = 1
        Do
            nbCol = nbCol + 1
        Loop Until wso.Cells(1, nbCol).Value = ""
        nbCol = nbCol - 1
        
        diffRow = diffRow + 1
        wsDiff.Cells(diffRow, 1).Value = wsName
        wsDiff.Cells(diffRow, 2).Value = nbCol
        
        Dim i As Integer
        For i = 1 To nbCol
            Set col1 = wso.Columns(i)
            Set col2 = wsn.Columns(i)
            readColumns = readColumns + 1
            
            'Compare Font Name
            If col1.Font.Name <> col2.Font.Name Then
                diffLog = diffLog & "Column " & i & " Font Name differs: " & col1.Font.Name & " vs " & col2.Font.Name & vbCrLf
                wsDiff.Cells(diffRow, 3).Value = i
                wsDiff.Cells(diffRow, 4).Value = col1.Font.Name
                wsDiff.Cells(diffRow, 5).Value = col2.Font.Name
            End If
            
            'Compare Font Size
            If col1.Font.size <> col2.Font.size Then
                diffLog = diffLog & "Column " & i & " Font Size differs: " & col1.Font.size & " vs " & col2.Font.size & vbCrLf
                wsDiff.Cells(diffRow, 3).Value = i
                wsDiff.Cells(diffRow, 4).Value = col1.Font.size
                wsDiff.Cells(diffRow, 5).Value = col2.Font.size
            End If
            
            'Compare Column Width
            If col1.ColumnWidth <> col2.ColumnWidth Then
                diffLog = diffLog & "Column " & i & " Width differs: " & col1.ColumnWidth & " vs " & col2.ColumnWidth & vbCrLf
                wsDiff.Cells(diffRow, 3).Value = i
                wsDiff.Cells(diffRow, 4).Value = col1.ColumnWidth
                wsDiff.Cells(diffRow, 5).Value = col2.ColumnWidth
            End If
            
            'Compare Number Format
            If col1.NumberFormat <> col2.NumberFormat Then
                diffLog = diffLog & "Column " & i & " Number Format differs: " & col1.NumberFormat & " vs " & col2.NumberFormat & vbCrLf
                wsDiff.Cells(diffRow, 3).Value = i
                wsDiff.Cells(diffRow, 4).Value = col1.NumberFormat
                wsDiff.Cells(diffRow, 5).Value = col2.NumberFormat
            End If
            
            'Compare Horizontal Alignment
            If col1.HorizontalAlignment <> col2.HorizontalAlignment Then
                diffLog = diffLog & "Column " & i & " Horizontal Alignment differs: " & col1.HorizontalAlignment & " vs " & col2.HorizontalAlignment & vbCrLf
                wsDiff.Cells(diffRow, 3).Value = i
                wsDiff.Cells(diffRow, 4).Value = col1.HorizontalAlignment
                wsDiff.Cells(diffRow, 5).Value = col2.HorizontalAlignment
            End If
    
            'Compare Background Color
            If col1.Interior.Color <> col2.Interior.Color Then
                diffLog = diffLog & "Column " & i & " Background Color differs: " & col1.Interior.Color & " vs " & col2.Interior.Color & vbCrLf
                wsDiff.Cells(diffRow, 3).Value = i
                wsDiff.Cells(diffRow, 4).Value = col1.Interior.Color
                wsDiff.Cells(diffRow, 5).Value = col2.Interior.Color
            End If
    
        Next i
        
    Next wso
    
    wsDiff.Columns.AutoFit
    wsDiff.Range("B:E").Columns.HorizontalAlignment = xlCenter
    
    'Result print setup - 2024-08-05 @ 05:16
    diffRow = diffRow + 2
    wsDiff.Range("A" & diffRow).Value = "**** " & Format$(readColumns, "###,##0") & _
                                        " colonnes analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:E" & diffRow)
    Call modAppli_Utils.AppliquerConditionalFormating(rngArea, 1, RGB(173, 216, 230))

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:E" & diffRow)
    Dim header1 As String: header1 = wb1.Name & " vs. " & wb2.Name
    Dim header2 As String: header2 = ""
    Call Simple_Print_Setup(wsDiff, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    'Output differences
    If diffLog <> "" Then
        MsgBox "Différences trouvées:" & vbCrLf & diffLog
    Else
        MsgBox "Aucune différence dans les colonnes."
    End If
    
    'Libérer la mémoire
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
    wsDiff.Range("A1").Value = "Worksheet"
    wsDiff.Range("B1").Value = "Prod_Cols"
    wsDiff.Range("C1").Value = "Dev_Cols"
    wsDiff.Range("D1").Value = "Prod_Rows"
    wsDiff.Range("E1").Value = "Dev_Rows"
    wsDiff.Range("F1").Value = "Ligne #"
    wsDiff.Range("G1").Value = "Colonne"
    wsDiff.Range("H1").Value = "Prod_Value"
    wsDiff.Range("I1").Value = "Dev_Value"
    Call Make_It_As_Header(wsDiff.Range("A1:I1"), RGB(0, 112, 192))

    'Set your workbooks and worksheets here
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\GCF_DataFiles\GCF_BD_MASTER_COPY.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx")
    
    Dim diffRow As Long
    diffRow = 1
    diffRow = diffRow + 1
    wsDiff.Cells(diffRow, 1).Value = "Prod: " & wb1.Name
    diffRow = diffRow + 1
    wsDiff.Cells(diffRow, 1).Value = "Dev : " & wb2.Name
    
    Dim wsProd As Worksheet
    Dim wsDev As Worksheet
    
    'Loop through each column (assuming both sheets have the same structure)
    Dim diffLogMess As String
    Dim readRows As Long
    Dim wsName As String
    For Each wsProd In wb1.Worksheets
        wsName = wsProd.Name
        Set wsDev = wb2.Sheets(wsName)
        
        'Determine number of columns and rows in Prod Workbook
        Dim arr(1 To 30) As String
        Dim nbColProd As Integer, nbRowProd As Long
        nbColProd = 0
        Do
            nbColProd = nbColProd + 1
            arr(nbColProd) = wsProd.Cells(1, nbColProd).Value
            Debug.Print "#044 - " & wsProd.Name, " Prod: ", wsProd.Cells(1, nbColProd).Value
        Loop Until wsProd.Cells(1, nbColProd).Value = ""
        nbColProd = nbColProd - 1
        nbRowProd = wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).Row
        
        'Determine number of columns and rows in Dev Workbook
        Dim nbColDev As Integer, nbRowDev As Long
        nbColDev = 0
        Do
            nbColDev = nbColDev + 1
            Debug.Print "#045 - " & wsDev.Name, " Dev : ", wsDev.Cells(1, nbColDev).Value
        Loop Until wsProd.Cells(1, nbColDev).Value = ""
        nbColDev = nbColDev - 1
        nbRowDev = wsDev.Cells(wsDev.Rows.count, 1).End(xlUp).Row
        
        diffRow = diffRow + 2
        wsDiff.Cells(diffRow, 1).Value = wsName
        wsDiff.Cells(diffRow, 2).Value = nbColProd
        wsDiff.Cells(diffRow, 3).Value = nbColDev
        wsDiff.Cells(diffRow, 4).Value = nbRowProd
        wsDiff.Cells(diffRow, 5).Value = nbRowDev
        
        Dim nbRow As Long
        If nbRowProd > nbRowDev Then
            wsDiff.Cells(diffRow, 6).Value = "Le client a ajouté " & nbRowProd - nbRowDev & " lignes dans la feuille"
            nbRow = nbRowProd
        End If
        If nbRowProd < nbRowDev Then
            wsDiff.Cells(diffRow, 6).Value = "Le dev a ajouté " & nbRowDev - nbRowProd & " lignes dans la feuille"
            nbRow = nbRowDev
        End If
        
        Dim rowProd As Range, rowDev As Range
        Dim i As Long, prevI As Long, j As Integer
        For i = 1 To nbRow
            Set rowProd = wsProd.Rows(i)
            Set rowDev = wsDev.Rows(i)
            readRows = readRows + 1
            
            For j = 1 To nbColProd
                If wsProd.Rows.Cells(i, j).Value <> wsDev.Rows.Cells(i, j).Value Then
                    diffLogMess = diffLogMess & "Cell(" & i & "," & j & ") was '" & _
                                  wsProd.Rows.Cells(i, j).Value & "' is now '" & _
                                  wsDev.Rows.Cells(i, j).Value & "'" & vbCrLf
                    diffRow = diffRow + 1
                    If i <> prevI Then
                        wsDiff.Cells(diffRow, 6).Value = "Ligne # " & i
                        prevI = i
                    End If
                    wsDiff.Cells(diffRow, 7).Value = j & "-" & arr(j)
                    wsDiff.Cells(diffRow, 8).Value = wsProd.Rows.Cells(i, j).Value
                    wsDiff.Cells(diffRow, 9).Value = wsDev.Rows.Cells(i, j).Value
                End If
            Next j
            
        Next i
        
    Next wsProd
    
    wsDiff.Columns.AutoFit
    wsDiff.Range("B:E").Columns.HorizontalAlignment = xlCenter
    wsDiff.Range("F:I").Columns.HorizontalAlignment = xlLeft
    
    'Result print setup - 2024-08-20 @ 05:48
    diffRow = diffRow + 2
    wsDiff.Range("A" & diffRow).Value = "**** " & Format$(readRows, "###,##0") & _
                                        " lignes analysées dans l'ensemble du Workbook ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:I" & diffRow)
    Call modAppli_Utils.AppliquerConditionalFormating(rngArea, 1, RGB(173, 216, 230))

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:I" & diffRow)
    Dim header1 As String: header1 = wb1.Name & " vs. " & wb2.Name
    Dim header2 As String: header2 = "Changements de lignes ou cellules"
    Call Simple_Print_Setup(wsDiff, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    'Output differences
    If diffLogMess <> "" Then
        MsgBox "Différences trouvées:" & vbCrLf & diffLogMess
    Else
        MsgBox "Aucune différence dans les lignes."
    End If
    
    'Libérer la mémoire
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

Sub Fix_Date_Format()
    
    'Initialisation de la boîte de dialogue FileDialog pour choisir le fichier Excel
    Dim fd As fileDialog
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
    'Configuration des filtres de fichiers (Excel uniquement)
    fd.Title = "Sélectionnez un fichier Excel"
    fd.Filters.Clear
    fd.Filters.Add "Fichiers Excel", "*.xlsx; *.xlsm"
    
    'Si l'utilisateur sélectionne un fichier, filePath contiendra son chemin
    Dim filePath As String
    Dim fileSelected As Boolean
    If fd.show = -1 Then
        filePath = fd.SelectedItems(1)
        fileSelected = True
    Else
        MsgBox "Aucun fichier sélectionné.", vbExclamation
        fileSelected = False
    End If
    
    'Ouvrir le fichier sélectionné s'il y en a un
    Dim wb As Workbook
    If fileSelected Then
        Set wb = Workbooks.Open(filePath)
        
        'Définir les colonnes spécifiques à nettoyer pour chaque feuille
        Dim colonnesANettoyer As Dictionary
        Set colonnesANettoyer = CreateObject("Scripting.Dictionary")
        
        colonnesANettoyer.Add "TEC_Local", Array("M") 'Vérifier et corriger la colonne D
        
        'Parcourir chaque feuille définie dans le dictionnaire
        Dim ws As Worksheet
        Dim cell As Range
        Dim dateOnly As Date
        Dim wsName As Variant
        Dim cols As Variant
        Dim col As Variant
        
        For Each wsName In colonnesANettoyer.keys
            'Vérifier si la feuille existe dans le classeur
            On Error Resume Next
            Set ws = wb.Sheets(wsName)
            Debug.Print "#046 - " & wsName
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                'Récupérer les colonnes à traiter pour cette feuille
                cols = colonnesANettoyer(wsName)
                
                'Parcourir chaque colonne spécifiée
                For Each col In cols
                    'Parcourir chaque cellule de la colonne spécifiée
                    For Each cell In ws.Columns(col).SpecialCells(xlCellTypeConstants)
                        'Vérifier si la cellule contient une date avec une heure
                        If IsDate(cell.Value) Then
                            'Vérifier si la valeur contient des heures (fraction décimale)
                            If cell.Value <> Int(cell.Value) Then
                                'Garde uniquement la partie date (sans heure)
                                Debug.Print "#047 - ", wsName & " - " & col & " - " & cell.Value
                                dateOnly = Int(cell.Value)
                                cell.Value = dateOnly
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
    
    'Libérer la mémoire
    Set cell = Nothing
    Set col = Nothing
    Set colonnesANettoyer = Nothing
    Set fd = Nothing
    Set wb = Nothing
    Set ws = Nothing
    Set wsName = Nothing
    
    MsgBox "Les dates ont été corrigées pour les colonnes spécifiques.", vbInformation

End Sub

Sub Debug_Écart_TEC_Local_vs_TEC_TDB_Data()

    Dim wsTEC As Worksheet: Set wsTEC = wsdTEC_Local
    Dim lurTEC As Long
    lurTEC = wsTEC.Cells(wsTEC.Rows.count, 1).End(xlUp).Row
    
    Dim wsTDB As Worksheet: Set wsTDB = wshTEC_TDB_Data
    Dim lurTDB As Long
    lurTDB = wsTDB.Cells(wsTDB.Rows.count, 1).End(xlUp).Row
    
    Dim wsOutput As Worksheet: Set wsOutput = wshzDocAnalyseEcartTEC
    Dim lastUsed As Long
    lastUsed = wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).Row + 2
    wsOutput.Range("A2:D" & lastUsed).ClearContents
    
    wsOutput.Cells(1, 1).Value = "TECID"
    wsOutput.Cells(1, 2).Value = "TEC_Local"
    wsOutput.Cells(1, 3).Value = "TEC_TDB_Data"
    wsOutput.Cells(1, 4).Value = "Vérification"
    
    Dim arr() As Variant
    ReDim arr(1 To 5000, 1 To 3)
    
    Dim i As Long
    Dim tecID As Long
    Dim dateCutOff As Date
    dateCutOff = Date
    
    Dim h As Currency, hTEC As Currency
    'Boucle dans TEC_Local
    Debug.Print "#048 - Mise en mémoire TEC_LOCAL"
    For i = 3 To lurTEC
        With wsTEC
            If .Range("D" & i).Value > dateCutOff Then Stop
            tecID = CLng(.Range("A" & i).Value)
            If arr(tecID, 1) <> "" Then Stop
            arr(tecID, 1) = tecID
            h = .Range("H" & i).Value
            If UCase$(.Range("N" & i).Value) = "VRAI" Then
                h = 0
            End If
            If h <> 0 Then
                If UCase$(.Range("J" & i).Value) = "VRAI" And Len(.Range("E" & i).Value) > 2 Then
                    If UCase$(.Range("L" & i).Value) = "FAUX" Then
                        If .Range("M" & i).Value <= dateCutOff Then
                            arr(tecID, 2) = h
                        Else
                            Stop
                        End If
                    End If
                End If
            End If
        End With
    Next i
    
    'Boucle dans TEC_TDB
    Dim hTDB As Double
    Debug.Print "#049 - Mise en mémoire TEC_TDB"
    For i = 2 To lurTDB
        With wsTDB
            If .Range("D" & i).Value > dateCutOff Then Stop
            tecID = CLng(.Range("A" & i).Value)
            arr(tecID, 1) = tecID
            arr(tecID, 3) = .Range("Q" & i).Value
        End With
    Next i
    
    Debug.Print "#050 - Analyse des écarts"
    Dim tTEC As Double, tTDB As Double
    Dim r As Long: r = 2
    wsOutput.Columns(2).EntireColumn.NumberFormat = "##0.00"
    wsOutput.Range("B:B").HorizontalAlignment = xlRight
    wsOutput.Columns(3).EntireColumn.NumberFormat = "##0.00"
    wsOutput.Range("C:C").HorizontalAlignment = xlRight
    
    For i = 1 To 5000
        tTEC = tTEC + arr(i, 2)
        tTDB = tTDB + arr(i, 3)
        If arr(i, 2) <> 0 Or arr(i, 3) <> 0 Then
            wsOutput.Cells(r, 1).Value = arr(i, 1)
            wsOutput.Cells(r, 2).Value = arr(i, 2)
            wsOutput.Cells(r, 3).Value = arr(i, 3)
            If arr(i, 2) <> arr(i, 3) Then
                wsOutput.Cells(r, 4).Value = "Valeurs sont différentes"
            End If
            r = r + 1
        End If
    Next i
    
    wsOutput.Cells(r + 1, 2).Value = Round(tTEC, 2)
    wsOutput.Cells(r + 1, 3).Value = Round(tTDB, 2)
    
    'Libérer la mémoire
    Set wsOutput = Nothing
    Set wsTEC = Nothing
    Set wsTDB = Nothing
    
    Debug.Print "#051 - Totaux", Round(tTEC, 2), Round(tTDB, 2)
    
End Sub

Sub Analyse_Search_For_Memory_Management()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("X_Doc_Search_Utility_Results")
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Cas")
    wsOutput.Range("A1").CurrentRegion.offset(1, 0).ClearContents
    
    Dim lastUsedRow As Long, r As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    r = 2
    
    Dim ligneCode As String, moduleName As String, procName As String
    Dim objetSet As String, objetForEach As String, objetNothing As String
    
    Dim added As String, cleared As String
    Dim i As Long
    For i = 2 To lastUsedRow
        If ws.Cells(i, 5).Value = "" Then
            Call SortDelimitedString(added, "|")
            Call SortDelimitedString(cleared, "|")
            If added <> cleared Then
                wsOutput.Cells(r, 1).Value = moduleName
                wsOutput.Cells(r, 2).Value = procName
                wsOutput.Cells(r, 3).Value = "'+ " & added
                wsOutput.Cells(r + 1, 3).Value = "'- " & cleared
                r = r + 3
            End If
            If ws.Cells(i + 1, 5).Value <> "" Then
                moduleName = ws.Cells(i + 1, 3).Value
                procName = ws.Cells(i + 1, 5).Value
            Else
                procName = ""
            End If
            added = ""
            cleared = ""
            GoTo Next_For
        End If
        ligneCode = Trim$(ws.Cells(i, 6))
        If InStr(ligneCode, "recSet As ") Then
            ligneCode = Replace(ligneCode, "recSet As ", "resste As ")
        End If
        If InStr(ligneCode, ".Recordset") Then
            ligneCode = Replace(ligneCode, ".Recordset", ".RecordSET")
        End If
        If InStr(ligneCode, ".Offset") Then
            ligneCode = Replace(ligneCode, ".Offset", ".offSET")
        End If
        If InStr(ligneCode, ".Offset") Then
            ligneCode = Replace(ligneCode, ".Offset", ".OffSET")
        End If
        
        objetSet = ""
        objetForEach = ""
        objetNothing = ""
        'Déclaration de l'objet avec Set...
        If InStr(ligneCode, "Set ") <> 0 Then
            If Left$(ligneCode, 4) = "Set " Or InStr(ligneCode, ": Set") <> 0 Then
                objetSet = Mid$(ligneCode, InStr(ligneCode, "Set ") + 4, Len(ligneCode))
                objetSet = Left$(objetSet, InStr(objetSet, " ") - 1)
                If objetSet = "As" Then Stop
                If InStr(added, objetSet & "|") = 0 Then
                    added = added + objetSet + "|"
                End If
            Else
                Debug.Print "#078 - " & ligneCode
            End If
        End If
        'Déclaration de l'objet avec For Each...
        If InStr(ligneCode, "For Each ") <> 0 Then
            objetForEach = Mid$(ligneCode, InStr(ligneCode, "For Each ") + 9, Len(ligneCode))
            objetForEach = Left$(objetForEach, InStr(objetForEach, " ") - 1)
            If objetForEach = "As" Then Stop
            If InStr(added, objetForEach & "|") = 0 Then
                added = added + objetForEach + "|"
            End If
        End If
        'Libération de l'objet avec = Nothing
        If InStr(ligneCode, " = Nothing") <> 0 Then
            objetNothing = Mid$(ligneCode, InStr(ligneCode, "Set") + 4, Len(ligneCode))
            objetNothing = Left$(objetNothing, InStr(objetNothing, " ") - 1)
            If objetNothing = "" Then Stop
            cleared = cleared + objetNothing + "|"
        End If
        
Next_For:
    Next i
    
    'Libérer la mémoire
    Set ws = Nothing
    Set wsOutput = Nothing
    
End Sub

Sub Sauvegarder_UserForms_Parameters() '2024-11-26 @ 07:42

    'Utiliser la feuille 'Doc_UserForm_Params' ou la créer pour sauvegarder les paramètres
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wshzDocUserFormParams
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Doc_UserForm_Params"
    End If
    On Error GoTo 0
    
    'En-têtes de colonnes
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Nom_UserForm", "Largeur", "Hauteur", "Position_Left", "Position_Top")
    
    Dim i As Integer
    i = 2
    'Parcourir tous les composants VBA pour trouver les UserForms
    Dim vbComp As Object
    Dim userFormName As String
    Dim uf As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = vbext_ct_MSForm Then
            userFormName = vbComp.Name
            On Error Resume Next
            ' Charger dynamiquement le UserForm
            Set uf = VBA.UserForms.Add(userFormName)
            On Error GoTo 0

            If Not uf Is Nothing Then
                ws.Cells(i, 1).Value = userFormName
                ws.Cells(i, 2).Value = uf.Width
                ws.Cells(i, 3).Value = uf.Height
                ws.Cells(i, 4).Value = uf.Left
                ws.Cells(i, 5).Value = uf.Top
                i = i + 1
                ' Décharger le UserForm pour libérer la mémoire
                Unload uf
                Set uf = Nothing
            End If
        End If
    Next vbComp
    
    'Libérer la mémoire
    Set uf = Nothing
    
    MsgBox "Paramètres des UserForms sauvegardés avec succès.", vbInformation

End Sub

Sub Restaurer_UserForms_Parameters()

    Dim ws As Worksheet
    Dim i As Integer
    Dim uf As Object
    Dim nomUF As String

    'Vérifier si la feuille existe
    On Error Resume Next
    Set ws = wshzDocUserFormParams
    If ws Is Nothing Then
        MsgBox "La feuille 'Doc_UserForm_Params' n'existe pas. Sauvegardez d'abord les paramètres.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    'Parcourir la liste des paramètres sauvegardés
    i = 2
    Do While ws.Cells(i, 1).Value <> ""
        nomUF = ws.Cells(i, 1).Value
        On Error Resume Next
        ' Charger dynamiquement le UserForm
        Set uf = VBA.UserForms.Add(nomUF)
        On Error GoTo 0

        If Not uf Is Nothing Then
            uf.Width = ws.Cells(i, 2).Value
            uf.Height = ws.Cells(i, 3).Value
            uf.Left = ws.Cells(i, 4).Value
            uf.Top = ws.Cells(i, 5).Value
            'Optionnel : afficher le UserForm pour vérifier
            'uf.Show
        End If

        i = i + 1
    Loop

    'Libérer la mémoire
    Set uf = Nothing
    Set ws = Nothing
    
    MsgBox "Paramètres des UserForms restaurés avec succès.", vbInformation

End Sub

Sub Get_UsedRange_In_Active_Workbook()

    Dim output As String
    
    'Feuille pour les résultats
    Dim feuilleNom As String
    feuilleNom = "X_Cellules_Utilisées"
    Call Erase_And_Create_Worksheet(feuilleNom)
    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Sheets(feuilleNom)
    Dim r As Long: r = 1
    wsOutput.Cells(r, 1).Value = "Feuille"
    wsOutput.Cells(r, 2).Value = "Plage utilisée"
    wsOutput.Cells(r, 3).Value = "Lignes utilisée"
    wsOutput.Cells(r, 4).Value = "Colonnes utilisée"
    wsOutput.Cells(r, 5).Value = "Nb. Cellules"
    r = r + 1
    
    'Parcourir chaque feuille du classeur
    Dim ws As Worksheet
    Dim cellCount As Long
    For Each ws In ThisWorkbook.Worksheets
        'Vérifier si UsedRange n'est pas vide
        On Error Resume Next
        Dim usedRange As Range
        Set usedRange = ws.usedRange
        On Error GoTo 0
        
        If Not usedRange Is Nothing Then
            ' Ajouter les informations à la sortie
            wsOutput.Cells(r, 1).Value = ws.Name
            wsOutput.Cells(r, 2).Value = usedRange.Address
            wsOutput.Cells(r, 3).Value = usedRange.Rows.count
            wsOutput.Cells(r, 4).Value = usedRange.Columns.count
            wsOutput.Cells(r, 5).Value = usedRange.Cells.count
        Else
            ' Si aucune cellule utilisée
            wsOutput.Cells(r, 1).Value = ws.Name
            wsOutput.Cells(r, 2).Value = "Aucune"
        End If
        r = r + 1
    Next ws
    
    MsgBox "Le traitement est complété. Voir la feuille '" & feuilleNom & "'", vbInformation
    
End Sub

Sub CreerRepertoireEtImporterFichiers() '2025-07-02 @ 13:57

    'Chemin du dossier contenant les fichiers PROD
    Dim cheminSourcePROD As String
    cheminSourcePROD = "P:\Administration\APP\GCF\DataFiles\"
    
    'Vérifier si des fichiers Actif_*.txt existent (utilisateurs encore présents)
    Dim actifFile As String
    Dim actifExists As Boolean
    actifFile = Dir(cheminSourcePROD & "Actif_*.txt")
    actifExists = (actifFile <> "")
    
    If actifExists Then
        MsgBox "Un ou plusieurs utilisateurs utilisent encore l'application." & vbNewLine & vbNewLine & _
               "La copie est annulée.", vbExclamation
        Exit Sub
    End If
    
    'Définir le chemin racine (local) pour la création du nouveau dossier
    Dim cheminRacineDestination As String
    cheminRacineDestination = "C:\VBA\GC_FISCALITÉ\GCF_DataFiles\"
    
    'Construire le nom du répertoire basé sur la date et l'heure actuelle
    Dim dateHeure As String
    Dim nouveauDossier As String
    dateHeure = Format$(Now, "yyyy_mm_dd_hhnn")
    nouveauDossier = cheminRacineDestination & dateHeure & "\"
    
    'Créer le répertoire s'il n'existe pas déjà (ne devrait pas exister)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.folderExists(nouveauDossier) Then
        fso.CreateFolder nouveauDossier
    End If
    
    'Noms des deux fichiers à copier (fixe)
    Dim nomFichier1 As String, nomFichier2 As String
    nomFichier1 = "GCF_BD_MASTER.xlsx"
    nomFichier2 = "GCF_BD_Entrée.xlsx"
    
    'Mise en place d'un fichier .lock chez le client - 2025-07-02 @ 14:03
    Dim fichierLock As String
    fichierLock = cheminSourcePROD & "GCF_BD_MASTER.lock"
    
    If fso.fileExists(fichierLock) Then
        MsgBox "Le fichier est déjà verrouillé sur l'environnement client." & vbNewLine & vbNewLine & _
               "L'import est annulée.", vbCritical
        Exit Sub
    Else
        'Écriture du fichier .lock avec infos utiles
        Dim fluxLock As Object
        Set fluxLock = fso.CreateTextFile(fichierLock, True)
        fluxLock.WriteLine "Fichier verrouillé par : Robert"
        fluxLock.WriteLine "Date : " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        fluxLock.WriteLine "Machine : " & Environ("COMPUTERNAME")
        fluxLock.Close
    End If
    
    'Copier le premier fichier
    If fso.fileExists(cheminSourcePROD & nomFichier1) Then
        fso.CopyFile Source:=cheminSourcePROD & nomFichier1, Destination:=nouveauDossier, OverwriteFiles:=False
    Else
        MsgBox "Fichier non trouvé : " & cheminSourcePROD & nomFichier1, vbExclamation, "Erreur"
    End If
    
    'Copier le deuxième fichier
    If fso.fileExists(cheminSourcePROD & nomFichier2) Then
        fso.CopyFile Source:=cheminSourcePROD & nomFichier2, Destination:=nouveauDossier, OverwriteFiles:=False
    Else
        MsgBox "Fichier non trouvé : " & cheminSourcePROD & nomFichier2, vbExclamation, "Erreur"
    End If

    'Copier les fichiers .log (variable)
    Dim fichier As String
    fichier = Dir(cheminSourcePROD & "*.log")
    Do While fichier <> ""
        'Copie du fichier PROD ---> Local
        fso.CopyFile Source:=cheminSourcePROD & fichier, Destination:=nouveauDossier, OverwriteFiles:=False
        'Efface le fichier PROD (initialiation)
        If fso.fileExists(cheminSourcePROD & fichier) Then Kill cheminSourcePROD & fichier
        'Fichier suivant à copier
        fichier = Dir
    Loop
    
    'Copie des deux fichiers du dossier temporaire vers le dossier DEV (but ultime)
    
    Dim dossierDEV As String
    dossierDEV = "C:\VBA\GC_FISCALITÉ\DataFiles\"
    
    'Copier le premier fichier
    If fso.fileExists(nouveauDossier & nomFichier1) Then
        fso.CopyFile Source:=cheminSourcePROD & nomFichier1, Destination:=dossierDEV, OverwriteFiles:=True
    Else
        MsgBox "Fichier non trouvé : " & nouveauDossier & nomFichier1, vbExclamation, "Erreur"
    End If
    
    'Copier le deuxième fichier
    If fso.fileExists(nouveauDossier & nomFichier2) Then
        fso.CopyFile Source:=cheminSourcePROD & nomFichier2, Destination:=dossierDEV, OverwriteFiles:=True
    Else
        MsgBox "Fichier non trouvé : " & nouveauDossier & nomFichier2, vbExclamation, "Erreur"
    End If

    MsgBox "Fichiers copiés dans le dossier : " & nouveauDossier, vbInformation, "Terminé"

End Sub

Sub AjusterEpurerTablesDeMaster() '2024-12-07 @ 06:47

    'Chemin du classeur à ajuster
    Dim cheminClasseur As String
    cheminClasseur = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx"

    'Ouvrir le classeur
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(cheminClasseur, ReadOnly:=False)
    If wb Is Nothing Then
        MsgBox "Impossible d'ouvrir le classeur 'GCF_BD_MASTER.xlsx'", vbExclamation, "Erreur"
        Exit Sub
    End If
    On Error GoTo 0

    '1. Supprimer les lignes facturées dans FAC_Projets_Détails et FAC_Projets_Entête - 2025-05-30 @ 07:17
    Dim i As Long
    Dim wsDetails As Worksheet, wsEntete As Worksheet

    'wsDetails et wsEntete du Workbook MASTER (pas les feuilles locales)
    Dim lastUsedRow As Long
    
    On Error Resume Next
    Set wsDetails = wb.Sheets("FAC_Projets_Détails")
    Set wsEntete = wb.Sheets("FAC_Projets_Entête")
    On Error GoTo 0

    If Not wsDetails Is Nothing Then
        With wsDetails
            lastUsedRow = .Cells(.Rows.count, "A").End(xlUp).Row
            If lastUsedRow >= 2 Then
                For i = lastUsedRow To 2 Step -1
                    If Trim(.Cells(i, "I").Value) = "-1" _
                       Or LCase(Trim(.Cells(i, "I").Value)) = "vrai" _
                       Or .Cells(i, "I").Value = True Then
                        .Rows(i).Delete
                    End If
                Next i
            End If
        End With
    End If

    If Not wsEntete Is Nothing Then
        With wsEntete
            lastUsedRow = .Cells(.Rows.count, "A").End(xlUp).Row
            If lastUsedRow >= 2 Then
                For i = lastUsedRow To 2 Step -1
                    If Trim(.Cells(i, "Z").Value) = "-1" _
                       Or LCase(Trim(.Cells(i, "Z").Value)) = "vrai" _
                       Or .Cells(i, "Z").Value = True Then
                        .Rows(i).Delete
                    End If
                Next i
            End If
        End With
    End If

    '2. Parcourir toutes les feuilles
    Dim ws As Worksheet
    Dim listeObjets As ListObjects
    Dim tableau As ListObject
    Dim DerniereLigne As Long
    Dim DerniereColonne As Long
    Dim nouvellePlage As Range
    
    For Each ws In wb.Worksheets
        Set listeObjets = ws.ListObjects
        'Parcourir chaque tableau de la feuille
        For Each tableau In listeObjets
            'Trouver la dernière ligne avec des données
            DerniereLigne = ws.Cells(ws.Rows.count, tableau.Range.Column).End(xlUp).Row
            'Trouver la dernière colonne avec des données
            DerniereColonne = ws.Cells(tableau.HeaderRowRange.row, ws.Columns.count).End(xlToLeft).Column
            'Redéfinir la plage du tableau
            Set nouvellePlage = ws.Range(ws.Cells(tableau.HeaderRowRange.row, tableau.Range.Column), _
                                         ws.Cells(DerniereLigne, DerniereColonne))
            On Error Resume Next
            tableau.Resize nouvellePlage
            On Error GoTo 0
        Next tableau
    Next ws

    '3. Enregistrer et fermer le classeur MASTER
    wb.Save
    wb.Close
    
    'Libérer la mémoire
    Set listeObjets = Nothing
    Set nouvellePlage = Nothing
    Set tableau = Nothing
    Set wb = Nothing
    Set wsDetails = Nothing
    Set wsEntete = Nothing
    
    MsgBox "Tous les tableaux ont été ajustés avec succès.", vbInformation, "Traitement est terminé"
    
End Sub

Sub VerifierControlesAssociesToutesFeuilles()

    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Sheets("Feuil2")
    wsOut.Range("A1").CurrentRegion.offset(1).Clear
    Dim r As Long
    
    Dim ws As Worksheet
    Dim shp As Shape
    Dim btn As Object
    Dim macroNameRaw As String
    Dim macroName As String
    Dim vbComp As Object
    Dim codeModule As Object
    Dim ligne As Long
    Dim found As Boolean
    Dim oleObj As OLEObject
    
    ' Parcourir toutes les feuilles du classeur
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print "#079 - Vérification des contrôles sur la feuille : " & ws.Name
        
        ' Vérification des Shapes (Formulaires ou Boutons assignés)
        For Each shp In ws.Shapes
            On Error Resume Next
            macroNameRaw = shp.OnAction
            On Error GoTo 0
            
            If macroNameRaw <> "" Then
                ' Extraire uniquement le nom de la macro après le "!"
                If InStr(1, macroNameRaw, "!") > 0 Then
                    macroName = Split(macroNameRaw, "!")(1)
                Else
                    macroName = macroNameRaw
                End If
                
                ' Vérifier si la macro existe
                found = VerifierMacroExiste(macroName)
                
                ' Résultat de la vérification
                r = r + 1
                wsOut.Cells(r, 1).Value = ws.Name
                wsOut.Cells(r, 2).Value = shp.Name
                wsOut.Cells(r, 3).Value = macroName
                wsOut.Cells(r, 4).Value = "shape"
                If found Then
                    wsOut.Cells(r, 5).Value = "Valide"
                Else
                    wsOut.Cells(r, 5).Value = "Manquante"
                End If
            End If
        Next shp
        
        ' Vérification des contrôles ActiveX
        For Each oleObj In ws.OLEObjects
            If TypeOf oleObj.Object Is MSForms.CommandButton Then
                ' Construire le nom de la macro à partir du nom du contrôle
                macroName = oleObj.Name & "_Click"
                
                ' Vérifier si la macro existe
                found = VerifierMacroExiste(macroName, ws.CodeName)
                
                ' Résultat de la vérification
                r = r + 1
                wsOut.Cells(r, 1).Value = ws.Name
                wsOut.Cells(r, 2).Value = oleObj.Name
                wsOut.Cells(r, 3).Value = macroName
                wsOut.Cells(r, 4).Value = "CommandButton"
                If found Then
                    wsOut.Cells(r, 5).Value = "Valide"
                Else
                    wsOut.Cells(r, 5).Value = "Manquante"
                End If
            End If
        Next oleObj
    Next ws

    wsOut.Activate
    
    MsgBox "Vérification terminée sur toutes les feuilles. Consultez la fenêtre Exécution pour les résultats.", vbInformation
    
End Sub

Function VerifierMacroExiste(macroName As String, Optional moduleName As String = "") As Boolean

    'Par defaut...
    VerifierMacroExiste = False
    
    'Si un module spécifique est fourni, vérifier uniquement dans ce module
    Dim vbComp As Object
    Dim codeModule As Object
    Dim ligne As Long
    
    If moduleName <> "" Then
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
        On Error GoTo 0
        If Not vbComp Is Nothing Then
            Set codeModule = vbComp.codeModule
            For ligne = 1 To codeModule.CountOfLines
                If codeModule.ProcOfLine(ligne, vbext_pk_Proc) = macroName Then
                    VerifierMacroExiste = True
                    Exit Function
                End If
            Next ligne
        End If
        Exit Function
    End If
    
    'Parcourir tous les modules si aucun module spécifique n'est fourni
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Set codeModule = vbComp.codeModule
        For ligne = 1 To codeModule.CountOfLines
            If codeModule.ProcOfLine(ligne, vbext_pk_Proc) = macroName Then
                VerifierMacroExiste = True
                Exit Function
            End If
        Next ligne
    Next vbComp
    
End Function

Sub main() '2024-12-25 @ 15:27

    'Feuille pour la sortie
    Dim outputName As String
    outputName = "Doc_File_Layouts"
    Call CreateOrReplaceWorksheet(outputName)
    
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets(outputName)
    
    'Tableau pour travailler en mémoire les résultats
    Dim outputArr() As String
    ReDim outputArr(1 To 500, 1 To 8)
    
    Dim outputRow As Long
    outputRow = 1
    
    Application.ScreenUpdating = False
    
    Call ListeEnumsGenerique("BD_Clients", 1, outputArr, outputRow)
    Call ListeEnumsGenerique("BD_Fournisseurs", 1, outputArr, outputRow)
    
    Call ListeEnumsGenerique("CC_Régularisations", 1, outputArr, outputRow)
    
    Call ListeEnumsGenerique("DEB_Récurrent", 1, outputArr, outputRow)
    Call ListeEnumsGenerique("DEB_Trans", 1, outputArr, outputRow)
    
    Call ListeEnumsGenerique("ENC_Détails", 1, outputArr, outputRow)
    Call ListeEnumsGenerique("ENC_Entête", 1, outputArr, outputRow)
    
    Call ListeEnumsGenerique("FAC_Comptes_Clients", 2, outputArr, outputRow)
    Call ListeEnumsGenerique("FAC_Détails", 2, outputArr, outputRow)
    Call ListeEnumsGenerique("FAC_Entête", 2, outputArr, outputRow)
    Call ListeEnumsGenerique("FAC_Projets_Détails", 1, outputArr, outputRow)
    Call ListeEnumsGenerique("FAC_Projets_Entête", 1, outputArr, outputRow)
    Call ListeEnumsGenerique("FAC_Sommaire_Taux", 1, outputArr, outputRow)
    
    Call ListeEnumsGenerique("GL_EJ_Récurrente", 1, outputArr, outputRow)
    Call ListeEnumsGenerique("GL_Trans", 1, outputArr, outputRow)
    
    Call ListeEnumsGenerique("TEC_Local", 2, outputArr, outputRow)
    Call ListeEnumsGenerique("TEC_TDB_Data", 1, outputArr, outputRow)
    
    Application.ScreenUpdating = True
    
    'Écriture des résultats (tableau) dans la feuille
    With wsOut
        .Cells.Clear 'Efface tout le contenu de la feuille
        .Range("A1").Resize(outputRow, UBound(outputArr, 2)).Value = outputArr
    End With
    
End Sub

Sub ListeEnumsGenerique(ByVal tableName As String, ByVal HeaderRow As Integer, ByRef arrArg() As String, ByRef outputRow As Long)

    'Obtenir la feuille de calcul
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(tableName)
    Dim saveTableName As String
    saveTableName = tableName
    
    Dim wb As Workbook
    If tableName = "BD_Clients" Or tableName = "BD_Fournisseurs" Then
        Set wb = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx")
        tableName = Replace(tableName, "BD_", "")
    Else
        Set wb = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx")
    End If
    Dim wsMaster As Worksheet
    If tableName <> "TEC_TDB_Data" Then
        Set wsMaster = wb.Sheets(tableName)
    End If
    tableName = saveTableName
    
    'Nom de la table
    arrArg(outputRow, 1) = tableName
    outputRow = outputRow + 1
    
    'Extraire la définition des Enum de la table à partir du code
    Dim arr() As Variant
    Call ExtractEnumDefinition(tableName, arr)
    
    'Boucle sur les colonnes
    Dim col As Long
    For col = LBound(arr, 1) To UBound(arr, 1)
        arrArg(outputRow, 1) = arr(col, 1)
        arrArg(outputRow, 2) = NumeroEnLettre(col)
        arrArg(outputRow, 3) = arr(col, 2)
        'Nom de la colonne dans la table
        arrArg(outputRow, 4) = ws.Cells(HeaderRow, col).Value
        If InStr(arr(col, 2), ws.Cells(HeaderRow, col).Value) = 0 Then
            arrArg(outputRow, 5) = "*"
        End If
        If Not wsMaster Is Nothing Then
            arrArg(outputRow, 6) = wsMaster.Cells(1, col).Value
            If InStr(arr(col, 2), wsMaster.Cells(1, col).Value) = 0 Then
                arrArg(outputRow, 7) = "*"
            End If
        End If
        'Valeurs des colonnes sur la première ligne de data
        arrArg(outputRow, 8) = ws.Cells(HeaderRow + 1, col).Value
        outputRow = outputRow + 1
    Next col
    
    'Ligne pour séparer les tables
    outputRow = outputRow + 1
    
    'Fermer sans sauvegarder
    wb.Close SaveChanges:=False
    
End Sub

Sub ExtractEnumDefinition(tableName As String, ByRef arr() As Variant)

    Dim LineNum As Long
    Dim TotalLines As Long
    Dim codeLine As String
    Dim InEnumBlock As Boolean
    Dim filePath As String
    
    'Variable de travail
    Dim EnumDefinition As String
    EnumDefinition = ""
    
    'Redimensionner le tableau
    ReDim arr(1 To 50, 1 To 2)
    Dim e As Long
    
    'Accéder au projet VBA actif
    Dim VBProj As VBIDE.VBProject
    Set VBProj = ThisWorkbook.VBProject

    'Parcourir tous les composants VBA
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.codeModule
    For Each vbComp In VBProj.VBComponents
        Set codeMod = vbComp.codeModule
        'Parcourir chaque ligne de code
        For LineNum = 1 To codeMod.CountOfLines
            codeLine = Trim$(codeMod.Lines(LineNum, 1))
            'Détection du début d'un Enum
            If InStr(1, codeLine, "Enum " & tableName, vbTextCompare) > 0 Then
                InEnumBlock = True
            ElseIf InEnumBlock Then
                'Détection de la fin de l'Enum
                If InStr(1, codeLine, "End Enum", vbTextCompare) > 0 Then
                    InEnumBlock = False
                    Exit For 'Terminer après l'extraction
                Else
                    'Ajouter les lignes à l'intérieur du Enum
                    If Left$(codeLine, 1) <> "[" Then
                        If Right$(codeLine, 11) = " = [_First]" Then
                            codeLine = Left$(codeLine, Len(codeLine) - 11)
                        End If
                        e = e + 1
                        arr(e, 1) = e
                        arr(e, 2) = codeLine
                        EnumDefinition = EnumDefinition & codeLine & "|"
                    End If
                End If
            End If
        Next LineNum
    Next vbComp

    'Redimension au minimum le tableau
    Call Array_2D_Resizer(arr, e, 2)
    
End Sub

Function CouleurEnRGBTableau(ByVal couleur As Long) As Variant

    Dim rgbArray(1 To 3) As Integer
    
    'Décomposer la couleur en composantes RGB
    rgbArray(1) = couleur Mod 256       ' Rouge
    rgbArray(2) = (couleur \ 256) Mod 256 ' Vert
    rgbArray(3) = (couleur \ 65536) Mod 256 ' Bleu
    
    'Retourner le tableau
    CouleurEnRGBTableau = rgbArray
    
End Function

Function Convertir_Couleur_RGB_Hex(ByVal couleur As Long) As String

    Dim rouge As Integer, vert As Integer, bleu As Integer
    
    ' Décomposer la couleur en composantes RGB
    rouge = couleur Mod 256
    vert = (couleur \ 256) Mod 256
    bleu = (couleur \ 65536) Mod 256
    
    'Construire la valeur HEX (en format #RRGGBB)
    Convertir_Couleur_RGB_Hex = "#" & Right$("00" & Hex$(rouge), 2) & _
                                        Right$("00" & Hex$(vert), 2) & _
                                        Right$("00" & Hex$(bleu), 2)
    
End Function

Sub test_CouleurEnRGBTableau()

    Dim couleur As Long
    Dim rgbArray As Variant
    
    wshMenuFAC.Activate
    wshMenuFAC.Range("A3").Select
    
    couleur = wshMenuFAC.Range("A3").Interior.Color
    couleur = gCOULEUR_BASE_FACTURATION
    
    rgbArray = CouleurEnRGBTableau(couleur)
    
    'Afficher les composantes
    MsgBox "Rouge: " & rgbArray(1) & ", Vert: " & rgbArray(2) & ", Bleu: " & rgbArray(3)
    
End Sub

Sub Test_Convertir_Couleur_RGB_Hex()

    Dim couleur As Long
    Dim couleurHex As String
    
    ' Obtenir la couleur de remplissage de la cellule
    couleur = 11854022
    
    ' Convertir en HEX
    couleurHex = Convertir_Couleur_RGB_Hex(couleur)
    
    ' Afficher le résultat
    MsgBox "La couleur HEX de la cellule A1 est : " & couleurHex
    
End Sub

Function Convertir_Couleur_OLE(ByVal couleur As Long) As String

    Dim rouge As Integer, vert As Integer, bleu As Integer
    
    'Décomposer la couleur en composantes RGB
    rouge = couleur Mod 256
    vert = (couleur \ 256) Mod 256
    bleu = (couleur \ 65536) Mod 256
    
    ' Construire le code OLE en inversant les composantes RGB en BGR
    Convertir_Couleur_OLE = "&H00" & Right$("00" & Hex$(bleu), 2) & _
                                        Right$("00" & Hex$(vert), 2) & _
                                        Right$("00" & Hex$(rouge), 2) & "&"
                                        
End Function

Sub Test_Convertir_Couleur_OLE()

    Dim couleur As Long
    Dim couleurOLE As String
    
    ' Exemple : couleur de la cellule A1
    couleur = gCOULEUR_BASE_FACTURATION
    
    ' Convertir en format OLE
    couleurOLE = Convertir_Couleur_OLE(couleur)
    
    ' Afficher la couleur en format OLE
    MsgBox "La couleur OLE est : " & couleurOLE
    
End Sub

Sub ValideNomProcedureCallLog()

    Dim ws As Worksheet
'    Set ws = Feuil5
    
    Dim lastUsedRow As Long
    lastUsedRow = 874
    
    Dim module As String, procedure As String, code As String
    Dim lineNo As Long
    Dim posPO As Integer, posPF As Integer, posCL As Integer
    Dim i As Integer
    For i = 2 To lastUsedRow
        module = ws.Range("C" & i).Value
        If module <> "" Then
            lineNo = ws.Range("D" & i).Value
            If lineNo = 325 Then Stop
            procedure = ws.Range("E" & i).Value
            procedure = Replace(procedure, "Sub ", "")
            procedure = Replace(procedure, "Function ", "")
            posPO = InStr(procedure, "(")
            posPF = InStr(procedure, ")")
            'Paramètres au complet sur la ligne -OU- Début seulement sur cette ligne
            If posPF > posPO Or (posPF = 0 And posPO <> 0) Then
                procedure = Trim$(Left$(procedure, posPO - 1))
                If InStr(procedure, "(") <> 0 Then Stop
            End If
            code = ws.Range("F" & i).Value
            posCL = InStr(code, "Call Log_Record")
            code = Mid$(code, posCL + 17)
            If InStr(code, module & ":" & procedure) = 0 Then
                Debug.Print i, module & ":" & procedure, code
            End If
        End If
    Next i
    
    MsgBox "Traitement terminé"
    
End Sub

Function NumeroEnLettre(ByVal num As Long) As String

    'Assurer que le nombre soit positif et supérieur à zéro
    If num <= 0 Then
        NumeroEnLettre = ""
        Exit Function
    End If
    
    'Construire la chaîne de caractères à partir du numéro
    Do
        num = num - 1
        NumeroEnLettre = Chr$(65 + (num Mod 26)) & NumeroEnLettre
        num = num \ 26
    Loop While num > 0
    
End Function

Sub ListerValidations()

    Dim ws As Worksheet
    Dim cell As Range
    Dim rngDV As Range
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    
    ' Vérifie s'il existe déjà une feuille de rapport, sinon la crée
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("ListeValidations")
    On Error GoTo 0
    
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "ListeValidations"
    Else
        ' Efface l'ancien contenu si la feuille existe déjà
        wsReport.Cells.Clear
    End If
    
    ' En-têtes de colonnes
    wsReport.Cells(1, 1).Value = "Feuille"
    wsReport.Cells(1, 2).Value = "Cellule"
    wsReport.Cells(1, 3).Value = "Type de Validation"
    wsReport.Cells(1, 4).Value = "Formule / Liste"
    
    rowIndex = 2

    ' Parcourt toutes les feuilles
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        ws.Unprotect
        Set rngDV = ws.Cells.SpecialCells(xlCellTypeAllValidation)
        On Error GoTo 0
        
        If Not rngDV Is Nothing Then
            For Each cell In rngDV
                With cell.Validation
                    wsReport.Cells(rowIndex, 1).Value = ws.Name
                    wsReport.Cells(rowIndex, 2).Value = cell.Address(False, False)
                    wsReport.Cells(rowIndex, 3).Value = .Type
                    If .Type = xlValidateList Then
                        wsReport.Cells(rowIndex, 4).Value = .Formula1 ' Affiche la liste ou la formule utilisée
                    Else
                        wsReport.Cells(rowIndex, 4).Value = "Autre type"
                    End If
                    rowIndex = rowIndex + 1
                End With
            Next cell
        End If
        
        Set rngDV = Nothing
    Next ws
    
    MsgBox "Liste des validations générée dans la feuille 'ListeValidations'.", vbInformation
    
End Sub

Sub AppliquerGrille(ws As Worksheet, plages As Variant)

    'Appliquer le grillage à chaque plage spécifiée
    Dim i As Integer
    For i = LBound(plages) To UBound(plages)
        Call CreerBorduresInterieures(ws.Range(plages(i)))
    Next i
    
End Sub

Sub CreerBorduresInterieures(rng As Variant) '2025-02-24 @ 16:40
    
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    'Appliquer les bordures intérieures (horizontales & verticales)
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
    End With

    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlHairline
    End With
            
End Sub

Sub DemarrerSauvegardeAutomatique() '2025-03-03 @ 07:19

    'Lancer l'export des modules VBA
    Call ExporterCodeVBA
    
    'Programmer la prochaine sauvegarde
    gNextBackupTime = Now + TimeValue("00:" & INTERVALLE_MINUTES_SAUVEGARDE & ":00")
    
    Application.OnTime gNextBackupTime, "DemarrerSauvegardeAutomatique"
    
End Sub

Sub StopperSauvegardeAutomatique()

    'Annuler la prochaine exécution prévue
    On Error Resume Next
    Application.OnTime gNextBackupTime, "DemarrerSauvegardeAutomatique", , False
    On Error GoTo 0
    
End Sub

Sub ExporterCodeVBA() '2025-03-11 @ 06:47

    'Définir le dossier où enregistrer les modules
    Dim dossierBackup As String
    dossierBackup = "C:\Users\RobertMV\OneDrive\_P E R S O N N E L\00_AU CAS OÙ\Backup_VBA\" & _
                            Format$(Now, "yyyy-mm-dd_HHMMSS") & "-" & ThisWorkbook.Name & "\"
    
    'Vérifier si le dossier existe, sinon le créer
    If Dir(dossierBackup, vbDirectory) = "" Then
        MkDir dossierBackup
    End If

    'Référence au projet VBA actif
    Dim ws As Workbook
    Set ws = ThisWorkbook

    'Parcourir tous les modules
    Dim vbComp As Object
    Dim ext As String
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: ext = ".bas" 'Module standard
            Case 2: ext = ".cls" 'Classe
            Case 3: ext = ".frm" 'UserForm
            Case vbext_ct_Document: ext = ".cls" 'Feuille de calcul et ThisWorkbook
            Case Else: ext = ""  'Autres (ignorés)
        End Select
        
        If ext <> "" Then
            vbComp.Export dossierBackup & vbComp.Name & ext
        End If
    Next vbComp

    'Libérer la mémoire
    Set vbComp = Nothing
    Set ws = Nothing
        
End Sub

Function VérifierAccesVBAAutorise() As Boolean

    Dim test As Object
    On Error Resume Next
    Set test = ThisWorkbook.VBProject.VBComponents
    VérifierAccesVBAAutorise = (Err.Number = 0)
    On Error GoTo 0
    
End Function

Sub Tester_dnrProf_Initials_Only() '2025-03-14 @ 10:42

    Dim nm As Name
    Dim rng As Range
    Dim strRef As String
    
    Set nm = ThisWorkbook.Names("dnrProf_Initials_Only")
    
    On Error Resume Next
    strRef = nm.RefersTo
    Set rng = Evaluate(strRef) 'Utiliser Evaluate pour contourner RefersToRange
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        MsgBox "Plage correcte : " & rng.Address
    Else
        MsgBox "Erreur : la plage nommée est invalide !", vbCritical
    End If
    
End Sub

Sub ComparerClasseursNiveauCellules()

    Dim wbOld As Workbook, wbNew As Workbook, wbReport As Workbook
    Dim wsOld As Worksheet, wsNew As Worksheet, wsReport As Worksheet
    Dim dictOld As Object, dictNew As Object
    Dim rngOld As Range, rngNew As Range
    Dim key As Variant, row As Range, lastRowOld As Long, lastRowNew As Long
    Dim reportRow As Long, col As Integer, lastCol As Integer
    Dim oldValues As Variant, newValues As Variant
    Dim diff As Boolean
    Dim fDialog As fileDialog
    
    'Sélection des fichiers
    Set fDialog = Application.fileDialog(msoFileDialogFilePicker)
    fDialog.Title = "Sélectionnez l'ancien classeur"
    If fDialog.show <> -1 Then Exit Sub
    Set wbOld = Workbooks.Open(fDialog.SelectedItems(1))
    
    fDialog.Title = "Sélectionnez le nouveau classeur"
    If fDialog.show <> -1 Then Exit Sub
    Set wbNew = Workbooks.Open(fDialog.SelectedItems(1))
    
    'Création du classeur de rapport
    Set wbReport = Workbooks.Add
    
    'Boucler sur les feuilles communes
    For Each wsOld In wbOld.Sheets
        On Error Resume Next
        Set wsNew = wbNew.Sheets(wsOld.Name)
        On Error GoTo 0
        
        If Not wsNew Is Nothing Then
            'Initialiser les dictionnaires
            Set dictOld = CreateObject("Scripting.Dictionary")
            Set dictNew = CreateObject("Scripting.Dictionary")
            
            'Déterminer la dernière ligne et colonne
            lastRowOld = wsOld.Cells(wsOld.Rows.count, 1).End(xlUp).Row
            lastRowNew = wsNew.Cells(wsNew.Rows.count, 1).End(xlUp).Row
            lastCol = wsOld.Cells(1, wsOld.Columns.count).End(xlToLeft).Column
            
            'Charger les données de l'ancien classeur
            Set rngOld = wsOld.Range("A2:A" & lastRowOld)
            For Each row In rngOld.Rows
                key = row.row & " - " & row.Cells(1, 1).Value & " " & row.Cells(1, 2).Value 'Clé unique (ajustez si nécessaire)
                dictOld(key) = row.EntireRow.Value
            Next row
            
            'Charger les données du nouveau classeur
            Set rngNew = wsNew.Range("A2:A" & lastRowNew)
            For Each row In rngNew.Rows
                key = row.row & " - " & row.Cells(1, 1).Value & " " & row.Cells(1, 2).Value
                dictNew(key) = row.EntireRow.Value
            Next row
            
            'Créer une feuille pour le rapport
            Set wsReport = wbReport.Sheets.Add
            wsReport.Name = "Diff " & wsOld.Name
            wsReport.Range("A1:D1").Value = Array("Élément", "Colonne", "Ancienne", "Nouvelle")
            reportRow = 2
            
            'Comparer les données cellule par cellule
            For Each key In dictOld.keys
                If Not dictNew.Exists(key) Then
                    'Ligne supprimée
                    wsReport.Cells(reportRow, 1).Value = key
                    wsReport.Cells(reportRow, 2).Value = "Ligne entière"
                    wsReport.Cells(reportRow, 3).Value = "Supprimée"
                    reportRow = reportRow + 1
                Else
                    'Vérifier chaque colonne individuellement
                    oldValues = dictOld(key)
                    newValues = dictNew(key)
                    For col = 1 To lastCol
                        If oldValues(1, col) <> newValues(1, col) Then
                            wsReport.Cells(reportRow, 1).Value = key
                            wsReport.Cells(reportRow, 2).Value = wsOld.Cells(1, col).Value 'Nom de la colonne
                            wsReport.Cells(reportRow, 3).Value = "Modifiée"
                            wsReport.Cells(reportRow, 4).Value = oldValues(1, col)
                            wsReport.Cells(reportRow, 5).Value = newValues(1, col)
                            reportRow = reportRow + 1
                        End If
                    Next col
                End If
            Next key
            
            'Vérifier les ajouts
            reportRow = reportRow + 1
            For Each key In dictNew.keys
                If Not dictOld.Exists(key) Then
                    wsReport.Cells(reportRow, 1).Value = key
                    wsReport.Cells(reportRow, 2).Value = "Ligne entière"
                    wsReport.Cells(reportRow, 3).Value = "Ajoutée"
                    reportRow = reportRow + 1
                End If
            Next key
        End If
    Next wsOld
    
    'Fermer les fichiers source sans enregistrer
    wbOld.Close False
    wbNew.Close False
    
    MsgBox "Comparaison terminée ! Consultez le classeur de rapport.", vbInformation
    
End Sub

Sub AnalyserImagesEntêteFactureExcel() '2025-05-27 @ 14:40

    Dim dossier As String, fichier As String
    Dim wb As Workbook, ws As Worksheet
    Dim img As Shape
    Dim largeurOrig As Double, hauteurOrig As Double
    Dim largeurActuelle As Double, hauteurActuelle As Double
    Dim cheminComplet As String
    Dim nomImageCible As String

    'Demande à l'utilisateur de choisir un dossier
    With Application.fileDialog(msoFileDialogFolderPicker)
        .Title = "Choisissez un dossier contenant les fichiers Excel"
        If .show <> -1 Then Exit Sub 'Annuler
        dossier = .SelectedItems(1)
    End With

    'Nom exact de l'image à trouver (ou utiliser un critère partiel)
    nomImageCible = "Image 1" '? Modifier si nécessaire

    'Recherche tous les fichiers .xlsx dans le dossier
    Dim dateSeuilMinimum As Date
    dateSeuilMinimum = DateSerial(2024, 8, 1)
    fichier = Dir(dossier & "\*.xlsx")

    Do While fichier <> ""
        cheminComplet = dossier & "\" & fichier
        If FileDateTime(cheminComplet) < dateSeuilMinimum Then
            fichier = Dir
            GoTo SkipFile
        End If
        Set wb = Workbooks.Open(cheminComplet, ReadOnly:=True)

        On Error Resume Next
        Set ws = wb.Worksheets(wb.Worksheets.count)
        If ws.Name = "Activités" Then
            GoTo SkipFile
        End If
        On Error GoTo 0

        If Not ws Is Nothing Then
            For Each img In ws.Shapes
                If img.Type = msoPicture Then
                    If img.Name = nomImageCible Then
                        largeurActuelle = img.Width
                        hauteurActuelle = img.Height

                        'Lire la taille originale estimée
                        Call LireTailleOriginaleImage(img, largeurOrig, hauteurOrig)

                        Debug.Print "Fichier : " & fichier
                        Debug.Print "  Image : " & img.Name
                        Debug.Print "  Taille actuelle : " & largeurActuelle & " x " & hauteurActuelle
                        Debug.Print "  Taille originale : " & largeurOrig & " x " & hauteurOrig
                        Debug.Print String(40, "-")
                    End If
                End If
            Next img
        End If

        wb.Close SaveChanges:=False
        fichier = Dir
SkipFile:
    Loop

    MsgBox "Analyse terminée."
    
End Sub

'Fonction pour estimer la taille originale d'une image
Sub LireTailleOriginaleImage(img As Shape, ByRef largeurOrig As Double, ByRef hauteurOrig As Double)

    Dim ws As Worksheet
    Dim copie As Shape

    Set ws = img.Parent
    img.Copy
    ws.Paste
    Set copie = ws.Shapes(ws.Shapes.count) 'la dernière collée

    With copie
        .ScaleWidth 1, msoTrue, msoScaleFromTopLeft
        .ScaleHeight 1, msoTrue, msoScaleFromTopLeft
        largeurOrig = .Width
        hauteurOrig = .Height
        .Delete
    End With
    
End Sub

'@Description "Compter le nombre de lignes dans le projet actif - 2025-06-18 @ 13:55"
Sub CompterLignesCode()

    Dim cheminComplet As String
    cheminComplet = "C:\Users\RobertMV\AppData\Roaming\Microsoft\AddIns\"
    
    Dim AddIn As String
    AddIn = "CompterLignesCodeProjet.xlam"
    
    Dim procedure As String
    procedure = "CompterLignesProjet"
    
    Call AppelerRoutineAddIn(cheminComplet & AddIn, procedure)
    
End Sub

'@Description "Appeler un AddIn - 2025-06-19 @ 06:54"
Sub AppelerRoutineAddIn(nomFichier As String, nomMacro As String)

    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks(nomFichier)
    On Error GoTo 0

    If wb Is Nothing Then
        If Dir(nomFichier) <> "" Then
            Set wb = Workbooks.Open(fileName:=nomFichier)
        Else
            MsgBox "Le fichier est introuvable : " & nomFichier, _
                   vbExclamation
            Exit Sub
        End If
    End If

    Application.Run "'" & nomFichier & "'!" & nomMacro
    
End Sub

Sub ScannerAppelsSubsImplicites() '2025-07-03 @ 17:49

    Dim comp As Object, codeMod As Object, dictSubs As Object
    Dim ligne As String, nomModule As String, lignes() As String
    Dim i As Long
    Dim nomProc As Variant

    Set dictSubs = BatirDictionnaireProcedures()

    Debug.Print "Appels implicites aux Sub sans 'Call'"

    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Or _
           comp.Type = vbext_ct_ClassModule Or _
           comp.Type = vbext_ct_MSForm Then

            nomModule = comp.Name
            lignes = Split(comp.codeModule.Lines(1, comp.codeModule.CountOfLines), vbCrLf)

            For i = 0 To UBound(lignes)
                ligne = Trim(lignes(i))
                'Ignore les commentaires, les lignes vides & les "Debug.Print"
                If ligne = "" Or Left(ligne, 1) = "'" Then GoTo LigneSuivante
                If Left(ligne, 12) = "Debug.Print " Then GoTo LigneSuivante
                If Left(ligne, 7) = "MsgBox " Then GoTo LigneSuivante
                For Each nomProc In dictSubs.keys
                    'Vérifie présence d'un nom de Sub et absence du mot 'Call' juste avant
                    If InStr(" " & ligne & " ", " " & nomProc & " ") > 0 And _
                        InStr(LCase(ligne), "call " & LCase(nomProc)) = 0 And _
                        InStr(LCase(ligne), "set " & LCase(nomProc)) = 0 Then
                        Debug.Print Pad(nomModule, 25) & " # " & Format(i + 1, "###0") & "   " & ligne
                    End If
                Next nomProc

LigneSuivante:
            Next i
        End If
    Next comp

    Debug.Print "Recherche terminée"
End Sub

Function BatirDictionnaireProcedures() As Object '2025-07-03 @ 17:53

    Dim comp As Object, codeMod As Object, dict As Object
    Dim ligne As String, nomSub As String
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Set codeMod = comp.codeModule

        For i = 1 To codeMod.CountOfLines
            ligne = Trim(codeMod.Lines(i, 1))
            'Ignore les commentaires ou les Functions
            If Left(ligne, 1) = "'" Then GoTo NextLigne
            If InStr(ligne, "Function") > 0 Then GoTo NextLigne

            If InStr(ligne, "Sub ") > 0 Then
                nomSub = codeMod.ProcOfLine(i, vbext_pk_Proc)
                If Not dict.Exists(nomSub) Then dict.Add nomSub, comp.Name
            End If

NextLigne:
        Next i
    Next comp

    Set BatirDictionnaireProcedures = dict
    
End Function

Function Pad(text As String, longueur As Integer) As String '2025-07-03 @ 17:54

    Pad = Left(text & Space(longueur), longueur)
    
End Function


