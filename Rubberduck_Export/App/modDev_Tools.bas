Attribute VB_Name = "modDev_Tools"
'@IgnoreModule UnassignedVariableUsage

Option Explicit

Sub ObtenirPlageAPartirDynamicNamedRange(dynamicRangeName As String, ByRef rng As Range)
    
    On Error Resume Next
    'Récupérer la formule associée au nom
    Dim refersToFormula As String
    refersToFormula = ThisWorkbook.Names(dynamicRangeName).RefersTo
    On Error GoTo 0
    
    If refersToFormula = vbNullString Then
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

Sub DetecterReferenceCirculaireDansClasseur() '2024-07-24 @ 07:31
    
    Dim circRef As String
    circRef = vbNullString
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

Sub zz_Comparer2ClasseursFormatColonnes() '2024-08-19 @ 16:24

    'Erase and create a new worksheet for differences
    Dim wsDiff As Worksheet
    Call CreerOuRemplacerFeuille("Différences_Colonnes")
    Set wsDiff = ThisWorkbook.Worksheets("Différences_Colonnes")
    wsDiff.Range("A1").Value = "Worksheet"
    wsDiff.Range("B1").Value = "Nb. colonnes"
    wsDiff.Range("C1").Value = "Colonne"
    wsDiff.Range("D1").Value = "Valeur originale"
    wsDiff.Range("E1").Value = "Nouvelle valeur"
    Call CreerEnteteDeFeuille(wsDiff.Range("A1:E1"), RGB(0, 112, 192))

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
        Loop Until wso.Cells(1, nbCol).Value = vbNullString
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
    Dim header2 As String: header2 = vbNullString
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsDiff, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    'Output differences
    If diffLog <> vbNullString Then
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

Sub zz_Comparer2ClasseursNiveauCellules() '2024-08-20 @ 05:14

    'Erase and create a new worksheet for differences
    Dim wsDiff As Worksheet
    Call CreerOuRemplacerFeuille("Différences_Lignes")
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
    Call CreerEnteteDeFeuille(wsDiff.Range("A1:I1"), RGB(0, 112, 192))

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
        Loop Until wsProd.Cells(1, nbColProd).Value = vbNullString
        nbColProd = nbColProd - 1
        nbRowProd = wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).Row
        
        'Determine number of columns and rows in Dev Workbook
        Dim nbColDev As Integer, nbRowDev As Long
        nbColDev = 0
        Do
            nbColDev = nbColDev + 1
            Debug.Print "#045 - " & wsDev.Name, " Dev : ", wsDev.Cells(1, nbColDev).Value
        Loop Until wsProd.Cells(1, nbColDev).Value = vbNullString
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
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsDiff, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    'Output differences
    If diffLogMess <> vbNullString Then
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

Sub zz_CorrigerFormatErroneDate()
    
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

Sub zz_ComparerValeursTECLocalVsTECTDBData()

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
            If arr(tecID, 1) <> vbNullString Then Stop
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

Sub zz_RechercherCodeVBAPourGestionMemoire()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("X_Doc_Search_Utility_Results")
    
    Dim wsOutput As Worksheet: Set wsOutput = wshzDocMemoryLeak
    wsOutput.Range("A1").CurrentRegion.offset(1, 0).ClearContents
    
    Dim lastUsedRow As Long, r As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    r = 2
    
    Dim ligneCode As String, moduleName As String, procName As String
    Dim objetSet As String, objetForEach As String, objetNothing As String
    
    Dim added As String, cleared As String
    Dim i As Long
    For i = 2 To lastUsedRow
        If ws.Cells(i, 5).Value = vbNullString Then
            Call TrierChaineAvecDelimiteurs(added, "|")
            Call TrierChaineAvecDelimiteurs(cleared, "|")
            If added <> cleared Then
                wsOutput.Cells(r, 1).Value = moduleName
                wsOutput.Cells(r, 2).Value = procName
                wsOutput.Cells(r, 3).Value = "'+ " & added
                wsOutput.Cells(r + 1, 3).Value = "'- " & cleared
                r = r + 3
            End If
            If ws.Cells(i + 1, 5).Value <> vbNullString Then
                moduleName = ws.Cells(i + 1, 3).Value
                procName = ws.Cells(i + 1, 5).Value
            Else
                procName = vbNullString
            End If
            added = vbNullString
            cleared = vbNullString
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
            ligneCode = Replace(ligneCode, ".Offset", ".Offset")
        End If
        If InStr(ligneCode, ".Offset") Then
            ligneCode = Replace(ligneCode, ".Offset", ".Offset")
        End If
        
        objetSet = vbNullString
        objetForEach = vbNullString
        objetNothing = vbNullString
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
            If objetNothing = vbNullString Then Stop
            cleared = cleared + objetNothing + "|"
        End If
        
Next_For:
    Next i
    
    'Libérer la mémoire
    Set ws = Nothing
    Set wsOutput = Nothing
    
End Sub

Sub CreerRepertoireEtImporterFichiers() '2025-07-02 @ 13:57

    'Chemin du dossier contenant les fichiers PROD
    Dim cheminSourcePROD As String
    cheminSourcePROD = "P:\Administration\APP\GCF\DataFiles\"
    
    'Vérifier si des fichiers Actif_*.txt existent (utilisateurs encore présents)
    Dim actifFile As String
    Dim actifExists As Boolean
    actifFile = Dir(cheminSourcePROD & "Actif_*.txt")
    actifExists = (actifFile <> vbNullString)
    
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
    nouveauDossier = cheminRacineDestination & dateHeure & Application.PathSeparator
    
    'Créer le répertoire s'il n'existe pas déjà (ne devrait pas exister)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.folderExists(nouveauDossier) Then
        fso.CreateFolder nouveauDossier
    End If
    
    'Noms des deux fichiers à copier (fixe)
    Dim nomFichier1 As String, nomFichier2 As String
    nomFichier1 = wsdADMIN.Range("MASTER_FILE").Value
    nomFichier2 = wsdADMIN.Range("CLIENTS_FILE").Value
    
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
        fso.CopyFile source:=cheminSourcePROD & nomFichier1, Destination:=nouveauDossier, OverwriteFiles:=False
    Else
        MsgBox "Fichier non trouvé : " & cheminSourcePROD & nomFichier1, vbExclamation, "Erreur"
    End If
    
    'Copier le deuxième fichier
    If fso.fileExists(cheminSourcePROD & nomFichier2) Then
        fso.CopyFile source:=cheminSourcePROD & nomFichier2, Destination:=nouveauDossier, OverwriteFiles:=False
    Else
        MsgBox "Fichier non trouvé : " & cheminSourcePROD & nomFichier2, vbExclamation, "Erreur"
    End If

    Dim fichier As String
    
    'Copier les fichiers .log (variable)
    fichier = Dir(cheminSourcePROD & "*.log")
    Do While fichier <> vbNullString
        'Copie du fichier PROD ---> Local
        fso.CopyFile source:=cheminSourcePROD & fichier, Destination:=nouveauDossier, OverwriteFiles:=False
        'Efface le fichier PROD (initialiation)
        If fso.fileExists(cheminSourcePROD & fichier) Then Kill cheminSourcePROD & fichier
        'Fichier suivant à copier
        fichier = Dir
    Loop
    
    'Copier les fichiers .txt (variable) '2025-07-11 @ 20:00
    fichier = Dir(cheminSourcePROD & "*.txt")
    Do While fichier <> vbNullString
        'Copie du fichier PROD ---> Local
        fso.CopyFile source:=cheminSourcePROD & fichier, Destination:=nouveauDossier, OverwriteFiles:=False
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
        fso.CopyFile source:=cheminSourcePROD & nomFichier1, Destination:=dossierDEV, OverwriteFiles:=True
    Else
        MsgBox "Fichier non trouvé : " & nouveauDossier & nomFichier1, vbExclamation, "Erreur"
    End If
    
    'Copier le deuxième fichier
    If fso.fileExists(nouveauDossier & nomFichier2) Then
        fso.CopyFile source:=cheminSourcePROD & nomFichier2, Destination:=dossierDEV, OverwriteFiles:=True
    Else
        MsgBox "Fichier non trouvé : " & nouveauDossier & nomFichier2, vbExclamation, "Erreur"
    End If

    MsgBox "Fichiers copiés dans le dossier : " & nouveauDossier, vbInformation, "Terminé"

End Sub


Sub shpSynchroniserDEVversPROD_Click()

    Call SynchroniserFichiers

End Sub

Sub SynchroniserFichiers() '2025-08-17 @ 18:43

    Dim cheminProd As String
    Dim cheminDev As String
    Dim fichierMaster As String
    Dim fichierEntree As String
    Dim fichierLock As String
    Dim dateModifDev As Date
    Dim dateModifProd As Date

    'Définir les chemins
    cheminProd = "P:\Administration\APP\GCF" & gDATA_PATH & Application.PathSeparator
    cheminDev = wsdADMIN.Range("PATH_DATA_FILES") & gDATA_PATH & Application.PathSeparator
    
    fichierLock = cheminProd & "GCF_BD_MASTER.lock"
    fichierMaster = "GCF_BD_MASTER.xlsx"
    fichierEntree = "GCF_BD_Entrée.xlsx"

    On Error GoTo GestionErreur

    'Étape 1 - Supprimer le fichier .lock
    If Dir(fichierLock) <> "" Then Kill fichierLock

    'Étape 2 - Copier GCF_BD_MASTER.xlsx si DEV est plus récent
    If Dir(cheminDev & fichierMaster) <> "" Then
        If Dir(cheminProd & fichierMaster) <> "" Then
            dateModifDev = FileDateTime(cheminDev & fichierMaster)
            dateModifProd = FileDateTime(cheminProd & fichierMaster)
            
            If dateModifDev > dateModifProd Then
                FileCopy cheminDev & fichierMaster, cheminProd & fichierMaster
            End If
        Else
            'Si le fichier n'existe pas en PROD, on le copie
            FileCopy cheminDev & fichierMaster, cheminProd & fichierMaster
        End If
    End If

    'Étape 3 - Copier GCF_BD_Entree.xlsx si DEV est plus récent
    If Dir(cheminDev & fichierEntree) <> "" Then
        If Dir(cheminProd & fichierEntree) <> "" Then
            dateModifDev = FileDateTime(cheminDev & fichierEntree)
            dateModifProd = FileDateTime(cheminProd & fichierEntree)
            
            If dateModifDev > dateModifProd Then
                FileCopy cheminDev & fichierEntree, cheminProd & fichierEntree
            End If
        Else
            'Si le fichier n'existe pas en PROD, on le copie
            FileCopy cheminDev & fichierEntree, cheminProd & fichierEntree
        End If
    End If

    MsgBox "Synchronisation terminée avec succès.", vbInformation
    Exit Sub

GestionErreur:
    MsgBox "SynchroniserFichiers - Erreur : " & Err.description, vbCritical
    
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

    '1. Supprimer les lignes facturées dans FAC_Projets_Details et FAC_Projets_Entete - 2025-05-30 @ 07:17
    Dim i As Long
    Dim wsDetails As Worksheet, wsEntete As Worksheet

    'wsDetails et wsEntete du Workbook MASTER (pas les feuilles locales)
    Dim lastUsedRow As Long
    
    On Error Resume Next
    Set wsDetails = wb.Sheets("FAC_Projets_Details")
    Set wsEntete = wb.Sheets("FAC_Projets_Entete")
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

Sub zz_CreerFileLayouts() '2024-12-25 @ 15:27

    'Feuille pour la sortie
    Dim outputName As String
    outputName = "Doc_File_Layouts"
    Call CreerOuRemplacerFeuille(outputName)
    
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets(outputName)
    
    'Tableau pour travailler en mémoire les résultats
    Dim outputArr() As String
    ReDim outputArr(1 To 500, 1 To 8)
    
    Dim outputRow As Long
    outputRow = 1
    
    Application.ScreenUpdating = False
    
    Call ListerEnumsGenerique("BD_Clients", 1, outputArr, outputRow)
    Call ListerEnumsGenerique("BD_Fournisseurs", 1, outputArr, outputRow)
    
    Call ListerEnumsGenerique("CC_Regularisations", 1, outputArr, outputRow)
    
    Call ListerEnumsGenerique("DEB_Recurrent", 1, outputArr, outputRow)
    Call ListerEnumsGenerique("DEB_Trans", 1, outputArr, outputRow)
    
    Call ListerEnumsGenerique("ENC_Details", 1, outputArr, outputRow)
    Call ListerEnumsGenerique("ENC_Entete", 1, outputArr, outputRow)
    
    Call ListerEnumsGenerique("FAC_Comptes_Clients", 2, outputArr, outputRow)
    Call ListerEnumsGenerique("FAC_Details", 2, outputArr, outputRow)
    Call ListerEnumsGenerique("FAC_Entete", 2, outputArr, outputRow)
    Call ListerEnumsGenerique("FAC_Projets_Details", 1, outputArr, outputRow)
    Call ListerEnumsGenerique("FAC_Projets_Entete", 1, outputArr, outputRow)
    Call ListerEnumsGenerique("FAC_Sommaire_Taux", 1, outputArr, outputRow)
    
    Call ListerEnumsGenerique("GL_EJ_Recurrente", 1, outputArr, outputRow)
    Call ListerEnumsGenerique("GL_Trans", 1, outputArr, outputRow)
    
    Call ListerEnumsGenerique("TEC_Local", 2, outputArr, outputRow)
    Call ListerEnumsGenerique("TEC_TDB_Data", 1, outputArr, outputRow)
    
    Application.ScreenUpdating = True
    
    'Écriture des résultats (tableau) dans la feuille
    With wsOut
        .Cells.Clear 'Efface tout le contenu de la feuille
        .Range("A1").Resize(outputRow, UBound(outputArr, 2)).Value = outputArr
    End With
    
End Sub

Sub ListerEnumsGenerique(ByRef tableName As String, ByVal HeaderRow As Integer, ByRef arrArg() As String, ByRef outputRow As Long)

    'Obtenir la feuille de calcul
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(tableName)
    Dim saveTableName As String
    saveTableName = tableName
    
    Dim wb As Workbook
    If tableName = "BD_Clients" Or tableName = "BD_Fournisseurs" Then
        Set wb = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx")
        tableName = Replace(tableName, "BD_", vbNullString)
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
    Call ExtraireEnumDefinition(tableName, arr)
    
    'Boucle sur les colonnes
    Dim col As Long
    For col = LBound(arr, 1) To UBound(arr, 1)
        arrArg(outputRow, 1) = arr(col, 1)
        arrArg(outputRow, 2) = Fn_ChiffreEnLettres(col)
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

Sub ExtraireEnumDefinition(tableName As String, ByRef arr() As Variant)

    Dim LineNum As Long
    Dim TotalLines As Long
    Dim codeLine As String
    Dim InEnumBlock As Boolean
    Dim filePath As String
    
    'Variable de travail
    Dim EnumDefinition As String
    EnumDefinition = vbNullString
    
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
    Call RedimensionnerTableau2D(arr, e, 2)
    
End Sub

Function Fn_CouleurEnRGBTableau(ByVal couleur As Long) As Variant

    Dim rgbArray(1 To 3) As Integer
    
    'Décomposer la couleur en composantes RGB
    rgbArray(1) = couleur Mod 256       ' Rouge
    rgbArray(2) = (couleur \ 256) Mod 256 ' Vert
    rgbArray(3) = (couleur \ 65536) Mod 256 ' Bleu
    
    'Retourner le tableau
    Fn_CouleurEnRGBTableau = rgbArray
    
End Function

Function Fn_ConvertirCouleurRGB2Hex(ByVal couleur As Long) As String

    Dim rouge As Integer, vert As Integer, bleu As Integer
    
    ' Décomposer la couleur en composantes RGB
    rouge = couleur Mod 256
    vert = (couleur \ 256) Mod 256
    bleu = (couleur \ 65536) Mod 256
    
    'Construire la valeur HEX (en format #RRGGBB)
    Fn_ConvertirCouleurRGB2Hex = "#" & Right$("00" & Hex$(rouge), 2) & _
                                        Right$("00" & Hex$(vert), 2) & _
                                        Right$("00" & Hex$(bleu), 2)
    
End Function

Sub zz_AfficheCouleurEnRGB()

    Dim couleur As Long
    Dim rgbArray As Variant
    
    wshMenuFAC.Activate
    wshMenuFAC.Range("A3").Select
    
    couleur = gCOULEUR_BASE_FACTURATION
    
    rgbArray = Fn_CouleurEnRGBTableau(couleur)
    
    'Afficher les composantes RGB
    MsgBox "Rouge: " & rgbArray(1) & ", Vert: " & rgbArray(2) & ", Bleu: " & rgbArray(3)
    
End Sub

Sub zz_ConvertirCouleurEnHEX()

    Dim couleur As Long
    Dim couleurHex As String
    
    'Obtenir la couleur de remplissage de la cellule
    couleur = 11854022
    
    'Convertir en HEX
    couleurHex = Fn_ConvertirCouleurRGB2Hex(couleur)
    
    'Afficher le résultat
    MsgBox "La couleur HEX de la couleur " & couleur & " est " & couleurHex
    
End Sub

Function Fn_ConvertiCouleurEnOLE(ByVal couleur As Long) As String

    Dim rouge As Integer, vert As Integer, bleu As Integer
    
    'Décomposer la couleur en composantes RGB
    rouge = couleur Mod 256
    vert = (couleur \ 256) Mod 256
    bleu = (couleur \ 65536) Mod 256
    
    ' Construire le code OLE en inversant les composantes RGB en BGR
    Fn_ConvertiCouleurEnOLE = "&H00" & Right$("00" & Hex$(bleu), 2) & _
                                        Right$("00" & Hex$(vert), 2) & _
                                        Right$("00" & Hex$(rouge), 2) & "&"
                                        
End Function

Sub zz_ConvertirCouleurOLE()

    Dim couleur As Long
    Dim couleurOLE As String
    
    'Exemple : couleur de la cellule A1
    couleur = gCOULEUR_BASE_FACTURATION
    
    'Convertir en format OLE
    couleurOLE = Fn_ConvertiCouleurEnOLE(couleur)
    
    'Afficher la couleur en format OLE
    MsgBox "La couleur OLE est : " & couleurOLE
    
End Sub

Function Fn_ChiffreEnLettres(ByVal num As Long) As String

    'Assurer que le nombre soit positif et supérieur à zéro
    If num <= 0 Then
        Fn_ChiffreEnLettres = vbNullString
        Exit Function
    End If
    
    'Construire la chaîne de caractères à partir du numéro
    Do
        num = num - 1
        Fn_ChiffreEnLettres = Chr$(65 + (num Mod 26)) & Fn_ChiffreEnLettres
        num = num \ 26
    Loop While num > 0
    
End Function

Sub zz_ListerValidations()

    Dim ws As Worksheet
    Dim cell As Range
    Dim rngDV As Range
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    
    'Vérifie s'il existe déjà une feuille de rapport, sinon la crée
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("DocListeValidations")
    On Error GoTo 0
    
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "DocListeValidations"
    Else
        'Efface l'ancien contenu si la feuille existe déjà
        wsReport.Cells.Clear
    End If
    
    'En-têtes de colonnes
    wsReport.Cells(1, 1).Value = "Feuille"
    wsReport.Cells(1, 2).Value = "Cellule"
    wsReport.Cells(1, 3).Value = "Type de Validation"
    wsReport.Cells(1, 4).Value = "Formule / Liste"
    
    rowIndex = 2

    'Parcourt de toutes les feuilles
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
                        wsReport.Cells(rowIndex, 4).Value = .Formula1 'Affiche la liste ou la formule utilisée
                    Else
                        wsReport.Cells(rowIndex, 4).Value = "Autre type"
                    End If
                    rowIndex = rowIndex + 1
                End With
            Next cell
        End If
        
        Set rngDV = Nothing
    Next ws
    
    MsgBox "Liste des validations générée dans la feuille 'DocListeValidations'.", vbInformation
    
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

Sub DemarrerSauvegardeCodeVBAAutomatique() '2025-03-03 @ 07:19

    'Lancer l'export des modules VBA
    Call ExporterCodeVBA
    
    'Programmer la prochaine sauvegarde
    gNextBackupTime = Now + TimeValue("00:" & INTERVALLE_MINUTES_SAUVEGARDE & ":00")
    
    Application.OnTime gNextBackupTime, "DemarrerSauvegardeCodeVBAAutomatique"
    
End Sub

Sub ArreterSauvegardeCodeVBA()

    'Annuler la prochaine exécution prévue
    On Error Resume Next
    Application.OnTime gNextBackupTime, "DemarrerSauvegardeCodeVBAAutomatique", , False
    On Error GoTo 0
    
End Sub

Sub ExporterCodeVBA() '2025-03-11 @ 06:47

    'Définir le dossier où enregistrer les modules
    Dim dossierBackup As String
    dossierBackup = "C:\Users\RobertMV\OneDrive\_P E R S O N N E L\00_AU CAS OÙ\Backup_VBA\" & _
                            Format$(Now, "yyyy-mm-dd_HHMMSS") & "-" & ThisWorkbook.Name & "\"
    
    'Vérifier si le dossier existe, sinon le créer
    If Dir(dossierBackup, vbDirectory) = vbNullString Then
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
            Case Else: ext = vbNullString  'Autres (ignorés)
        End Select
        
        If ext <> vbNullString Then
            vbComp.Export dossierBackup & vbComp.Name & ext
        End If
    Next vbComp

    'Libérer la mémoire
    Set vbComp = Nothing
    Set ws = Nothing
        
End Sub

Sub zz_ComparerClasseursNiveauCellules()

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

'@Description ("Compter le nombre de lignes dans le projet actif")
Sub CompterLignesCode() '2025-06-18 @ 13:55
Attribute CompterLignesCode.VB_Description = "Compter le nombre de lignes dans le projet actif"

    Dim cheminComplet As String
    cheminComplet = "C:\Users\RobertMV\AppData\Roaming\Microsoft\AddIns\"
    
    Dim AddIn As String
    AddIn = "CompterLignesCodeProjet.xlam"
    
    Dim procedure As String
    procedure = "CompterLignesProjet"
    
    Call AppelerRoutineAddIn(cheminComplet & AddIn, procedure)
    
End Sub

'@Description "Appeler un AddIn"
Sub AppelerRoutineAddIn(nomFichier As String, nomMacro As String) '2025-06-19 @ 06:54

    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks(nomFichier)
    On Error GoTo 0

    If wb Is Nothing Then
        If Dir(nomFichier) <> vbNullString Then
            Set wb = Workbooks.Open(fileName:=nomFichier)
        Else
            MsgBox "Le fichier est introuvable : " & nomFichier, _
                   vbExclamation
            Exit Sub
        End If
    End If

    Application.Run "'" & nomFichier & "'!" & nomMacro
    
End Sub

Sub zz_ObtenirListeAppelSubsSansCall() '2025-08-05 @ 13:44

    Dim comp As Object
    Dim codeMod As Object
    Dim dictSubs As Object
    
    Set dictSubs = Fn_BatirDictionnaireProcedures()

    Debug.Print "Liste des appels aux Subs SANS 'Call'"

    Dim ligne As String
    Dim nomModule As String
    Dim lignes() As String
    Dim i As Long
    Dim nomProc As Variant
    Dim cas As Long
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Or _
           comp.Type = vbext_ct_ClassModule Or _
           comp.Type = vbext_ct_MSForm Then

            nomModule = comp.Name
            lignes = Split(comp.codeModule.Lines(1, comp.codeModule.CountOfLines), vbCrLf)

            For i = 0 To UBound(lignes)
                ligne = Trim(lignes(i))
                'Ignore les commentaires, les lignes vides & les "Debug.Print"
                If ligne = vbNullString Or Left(ligne, 1) = "'" Then GoTo LigneSuivante
                If Left(ligne, 12) = "Debug.Print " Then GoTo LigneSuivante
                If Left(ligne, 7) = "MsgBox " Then GoTo LigneSuivante
                For Each nomProc In dictSubs.keys
                    'Vérifie présence d'un nom de Sub et absence du mot 'Call' juste avant
                    If InStr(" " & ligne & " ", " " & nomProc & " ") > 0 And _
                        InStr(LCase(ligne), "call " & LCase(nomProc)) = 0 And _
                        InStr(LCase(ligne), "set " & LCase(nomProc)) = 0 Then
                        Debug.Print Fn_PadDroite(nomModule, 25) & " # " & Format(i + 1, "###0") & "   " & ligne
                        cas = cas + 1
                    End If
                    
                Next nomProc

LigneSuivante:
            Next i
        End If
    Next comp
    
    'Libérer la mémoire
    Set codeMod = Nothing
    Set comp = Nothing
    Set dictSubs = Nothing
  
    'Recherche terminée
    If cas = 0 Then
        Debug.Print "Recherche terminée, sans aucun cas"
    Else
        Debug.Print "Recherche terminée, avec " & cas & " cas d'appel sans 'Call'"
    End If
    
End Sub

Function Fn_BatirDictionnaireProcedures() As Object '2025-07-03 @ 17:53

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
            If InStr(ligne, "Function ") > 0 Then GoTo NextLigne

            If InStr(ligne, "Sub ") > 0 Then
                nomSub = codeMod.ProcOfLine(i, vbext_pk_Proc)
                If Not dict.Exists(nomSub) Then dict.Add nomSub, comp.Name
            End If

NextLigne:
        Next i
    Next comp

    Set Fn_BatirDictionnaireProcedures = dict
    
End Function

Function Fn_PadDroite(text As String, longueur As Integer) As String '2025-07-03 @ 17:54

    Fn_PadDroite = Left(text & Space(longueur), longueur)
    
End Function

Sub zz_InventaireProceduresEtFonctions() '2025-08-11 @ 10:54

    Dim comp As VBIDE.VBComponent
    Dim i As Long, ligne As String
    Dim regexDecl As Object, regexCall As Object
    Dim tableau()
    Dim r As Long

    ' Initialiser expressions régulières
    Set regexDecl = CreateObject("VBScript.RegExp")
    With regexDecl
        .pattern = "^\s*(Public|Private)?\s*(Sub|Function|Property\s+(Get|Let|Set))"
        .IgnoreCase = True
        .Global = False
    End With

    Set regexCall = CreateObject("VBScript.RegExp")
    With regexCall
        .pattern = "(OnAction\s*=\s*""[^""]+""|Application\.Run\s*""[^""]+""|CallByName\s*\(.*?""[^""]+""[^)]*\))"
        .IgnoreCase = True
        .Global = True
    End With

    ' Dictionnaire des noms déclarés
    Dim dictNomModule As Object: Set dictNomModule = CreateObject("Scripting.Dictionary")
    ReDim tableau(1 To 2000, 1 To 9)
    r = 0

    ' Parcours des composants VBA
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Dim typeModule As String
        Select Case comp.Type
        Case vbext_ct_StdModule
            typeModule = "3_Standard"
        Case vbext_ct_ClassModule
            typeModule = "4_Classe"
        Case vbext_ct_Document
            typeModule = "1_Feuille/Workbook"
        Case vbext_ct_MSForm
            typeModule = "2_UserForm"
        Case Else
            typeModule = "z_Autre"
        End Select
        For i = 1 To comp.codeModule.CountOfLines
            ligne = comp.codeModule.Lines(i, 1)

            ' Déclaration
            If regexDecl.test(ligne) Then
                r = r + 1
                Dim nomProc As String: nomProc = Fn_ExtraireNomProcedure(ligne)
                Dim t As String, p As String
                Call Fn_ExtraireTypeEtPortee(ligne, t, p)

                tableau(r, 1) = comp.Name
                tableau(r, 2) = nomProc
                tableau(r, 3) = t
                tableau(r, 4) = p
                tableau(r, 5) = ligne
                tableau(r, 6) = "Déclaration"
                tableau(r, 7) = typeModule
                tableau(r, 8) = i
                tableau(r, 9) = ""

                If nomProc <> "" Then
                    If Not dictNomModule.Exists(nomProc) Then
                        dictNomModule.Add nomProc, comp.Name
                    Else
                        dictNomModule(nomProc) = dictNomModule(nomProc) & " | " & comp.Name
                    End If
                End If
            End If

            ' Appels indirects
            If regexCall.test(ligne) Then
                Dim matchesCall: Set matchesCall = regexCall.Execute(ligne)
                Dim appel
                For Each appel In matchesCall
                    r = r + 1
                    Dim nomAppel As String: nomAppel = Fn_NomProcedureIndirect(appel.Value)

                    tableau(r, 1) = comp.Name
                    tableau(r, 2) = nomAppel
                    tableau(r, 3) = "---"
                    tableau(r, 4) = "---"
                    tableau(r, 5) = ligne
                    If InStr(appel, "OnAction") > 0 Then
                        tableau(r, 6) = "Appel indirect (.OnAction)"
                    ElseIf InStr(appel, "Application.Run") > 0 Then
                        tableau(r, 6) = "Appel indirect (Application.Run)"
                    ElseIf InStr(appel, "CallByName") > 0 Then
                        tableau(r, 6) = "Appel indirect (CallByName)"
                    Else
                        tableau(r, 6) = "Appel indirect (autre)"
                    End If
                    tableau(r, 7) = typeModule
                    tableau(r, 8) = i

                    If nomAppel <> "" Then
                        If dictNomModule.Exists(nomAppel) Then
                            tableau(r, 9) = ""
                        Else
                            tableau(r, 9) = "Non trouvé"
                        End If
                    Else
                        tableau(r, 9) = vbNullString
                    End If
                Next appel
            End If
        Next i
    Next comp

    'Écriture dans la feuille
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("InventaireProcedures").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "InventaireProcedures"
    ws.Range("A1:I1").Value = Array("Module", "Nom", "Type", "Portée", "Contenu", "Catégorie", "Type de module", "Ligne", "Vérification")
    ws.Range("A2").Resize(r, 9).Value = tableau

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("G2:G" & r), Order:=xlAscending ' Type de module
        .SortFields.Add key:=ws.Range("A2:A" & r), Order:=xlAscending ' Module
        .SortFields.Add key:=ws.Range("B2:B" & r), Order:=xlAscending ' Nom
        .SetRange ws.Range("A2:I" & r)
        .Header = xlNo
        .Apply
    End With
    
    With ws.Range("A1:I1")
        .Interior.Color = RGB(180, 198, 231)     ' Bleu pâle
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Columns.AutoFit
    ws.Range("A1:I1").AutoFilter
    
    With ws.Range("A2:I" & r)
        For i = 2 To r
            If (i Mod 2 = 0) Then
                ws.Range("A" & i & ":I" & i).Interior.Color = RGB(240, 240, 240)
            End If
        Next i
    End With

    With ws.PageSetup
        .Orientation = xlLandscape
        .TopMargin = Application.CentimetersToPoints(0.4)
        .BottomMargin = Application.CentimetersToPoints(0.4)
        .LeftMargin = Application.CentimetersToPoints(0.4)
        .RightMargin = Application.CentimetersToPoints(0.4)
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .PrintTitleRows = "$1:$1"
        .LeftFooter = Format(Now, "yyyy-mm-dd à HH:mm:ss")
        .RightFooter = "Page &P sur &N"
    End With

    MsgBox "Analyse terminée : " & r & " entrées listées dans 'InventaireProcedures'.", vbInformation
    
End Sub

Function Fn_ExtraireNomProcedure(ligne As String) As String '2025-07-15 @ 22:56

    Dim mots() As String, i As Long
    ligne = Trim(ligne)
    If InStr(ligne, "(") > 0 Then ligne = Left(ligne, InStr(ligne, "(") - 1)
    ligne = Replace(ligne, ":", " ")
    mots = Split(ligne)
    For i = UBound(mots) To 0 Step -1
        If mots(i) <> "" Then
            Fn_ExtraireNomProcedure = mots(i)
            Exit Function
        End If
    Next i
    
End Function

Function Fn_NomProcedureIndirect(texte As String) As String '2025-07-15 @ 22:56

    Dim debut As Long, fin As Long
    Fn_NomProcedureIndirect = ""
    debut = InStr(texte, """")
    If debut > 0 Then
        fin = InStr(debut + 1, texte, """")
        If fin > debut Then
            Fn_NomProcedureIndirect = Mid(texte, debut + 1, fin - debut - 1)
        End If
    End If
    
End Function

Function Fn_ExtraireTypeEtPortee(ligne As String, ByRef TypeRetour As String, ByRef PorteeRetour As String) '2025-07-15 @ 22:56

    Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
    reg.pattern = "^\s*(Public|Private)?\s*(Sub|Function|Property\s+(Get|Let|Set))"
    reg.IgnoreCase = True
    reg.Global = False

    If reg.test(ligne) Then
        Dim matches: Set matches = reg.Execute(ligne)
        PorteeRetour = matches(0).SubMatches(0)
        If PorteeRetour = "" Then PorteeRetour = "Public"

        If matches(0).SubMatches(1) = "Property" Then
            TypeRetour = "Property " & matches(0).SubMatches(2)
        Else
            TypeRetour = matches(0).SubMatches(1)
        End If
    Else
        TypeRetour = "---": PorteeRetour = "---"
    End If
    
End Function

Function Fn_ConstruireDictionnaireDeclarations() As Object '2025-07-15 @ 22:56

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = Worksheets("InventaireProcedures")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    Dim i As Long, nom As String, moduleNom As String
    For i = 2 To lastRow
        If ws.Cells(i, 6).Value = "Déclaration" Then
            nom = Trim(ws.Cells(i, 2).Value)
            moduleNom = Trim(ws.Cells(i, 1).Value)
            If nom <> "" Then
                If Not dict.Exists(nom) Then
                    dict.Add nom, moduleNom
                End If
            End If
        End If
    Next i

    Set Fn_ConstruireDictionnaireDeclarations = dict
    
End Function

Sub zz_InjecterModuleDansAppels() '2025-07-15 @ 22:56

    Dim dict As Object: Set dict = Fn_ConstruireDictionnaireDeclarations()
    Dim ws As Worksheet: Set ws = Worksheets("InventaireProcedures")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    Dim i As Long, nomAppel As String, moduleAppelant As String, moduleCible As String

    For i = 2 To lastRow
        If ws.Cells(i, 6).Value Like "Appel indirect*" Then
            nomAppel = Trim(ws.Cells(i, 2).Value)
            moduleAppelant = Trim(ws.Cells(i, 1).Value)

            If dict.Exists(nomAppel) Then
                moduleCible = dict(nomAppel)
                If moduleAppelant <> moduleCible Then
                    Debug.Print "Appel externe : '" & nomAppel & "' devient '" & moduleCible & "." & nomAppel & "'"
                    'Option : marquer dans colonne 10
                    ws.Cells(i, 10).Value = moduleCible & "." & nomAppel
'                Else
'                    Debug.Print "Appel interne : laisser """ & nomAppel & """"
                End If
            Else
                Debug.Print "Nom non reconnu (" & i & ") '" & nomAppel & "'"
            End If
        End If
    Next i
    
End Sub

Sub zz_VerifierCombinaisonClientIDClientNomDansTEC()

    'Fichier maître des clients
    Dim strFile As String
    strFile = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
    Dim wb As Workbook
    Set wb = Workbooks.Open(strFile)
    
    'Feuille TEC_Local
    Dim ws As Worksheet
    Set ws = wsdTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    'Transfère la feuille en mémoire (matrice)
    Dim m As Variant
    m = ws.Range("A3:P" & lastUsedRow).Value
    
    'Feuille de travail (ouput)
    Dim output As Worksheet
    Set output = ThisWorkbook.Sheets("Feuil1")
    Dim r As Integer
    r = 1
    output.Cells.Clear
    output.Cells(r, 1) = "Ligne"
    output.Cells(r, 2) = "TEC_ID"
    output.Cells(r, 3) = "ClientID"
    output.Cells(r, 4) = "ClientName"
    output.Cells(r, 5) = "clientNameFromMF"
    output.Cells(r, 6) = "Date"
    output.Cells(r, 7) = "Prof"
    output.Cells(r, 8) = "Description"
    output.Cells(r, 9) = "Heures"
    output.Cells(r, 10) = "estFacturée"
    
    Dim clientID As String, clientName As String, clientNameFromMF As String
    Dim allCols As Variant
    Dim i As Integer
    For i = 1 To UBound(m, 1)
        clientID = m(i, fTECClientID)
        clientName = m(i, fTECClientNom)
        
        'Obtenir le nom du client associé à clientID
        allCols = Fn_ObtenirLigneDeFeuille("BD_Clients", clientID, fClntFMClientID)
        'Vérifier le résultat retourné
        If IsArray(allCols) Then
            clientNameFromMF = allCols(1)
        Else
            MsgBox "Valeur non trouvée !!!", vbCritical
        End If
        
        If clientName <> clientNameFromMF Then
            r = r + 1
            output.Cells(r, 1).Value = i + 2
            output.Cells(r, 2).Value = m(i, fTECTECID)
            output.Cells(r, 3).Value = clientID
            output.Cells(r, 4).Value = clientName
            output.Cells(r, 5).Value = clientNameFromMF
            output.Cells(r, 6).Value = m(i, fTECDate)
            output.Cells(r, 7).Value = m(i, fTECProf)
            output.Cells(r, 8).Value = m(i, fTECDescription)
            output.Cells(r, 9).Value = m(i, fTECHeures)
            output.Cells(r, 10).Value = m(i, fTECEstFacturee)
        End If
        
    Next i

    Debug.Print lastUsedRow, UBound(m, 1)

    wb.Close False
    
End Sub

Sub zz_FixEstFacturable() '2025-07-23 @ 08:26

    Dim cheminFichier As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim valeurActuelle As Variant
    
    'Chemin d’accès au classeur cible
    cheminFichier = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Ouvrir le classeur
    Set wb = Workbooks.Open(fileName:=cheminFichier, ReadOnly:=False)
    Set ws = wb.Sheets("TEC_Local")

    'Parcourir les lignes
    For i = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).Row
        If Len(Trim(ws.Cells(i, fTECClientID).Value)) < 2 And _
               ws.Cells(i, fTECEstFacturable).Value = "VRAI" Then
            valeurActuelle = ws.Cells(i, fTECEstFacturable).Value
            ws.Cells(i, fTECEstFacturable) = "FAUX"
        End If
    Next i

    'Sauvegarder les modifications et fermer le classeur
    wb.Save
    wb.Close SaveChanges:=True

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Le traitement est complété" & vbNewLine & vbNewLine & _
            "La colonne 'estFacturable' est en ligne avec le client", _
            vbInformation

End Sub

Public Function ListerPDFs(dossier As String) As Object '2025-07-23 @ 12:40

    Dim dictPDFs As Object
    Set dictPDFs = CreateObject("Scripting.Dictionary")

    Dim facturesManquantesConnues As String
    facturesManquantesConnues = "24-24540A.24-24540B.24-24548A.24-24548B.24-24552v2.24-24566v2." & _
        "24-24655 v2.24-24721 V2.25-24756A.25-24756B.25-24761v2.25-24937V2.25-25074A.25-25074B." & _
        "25-25078-25-25107.99-25046.Facture #23466.Facture #24059.Facture #24133.Facture #24224." & _
        "Facture #24324."
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dossierObj As Object
    Set dossierObj = fso.GetFolder(dossier)

    Dim fichier As Object
    Dim nomSansExtension As String
    For Each fichier In dossierObj.Files
        If LCase(fso.GetExtensionName(fichier.Name)) = "pdf" Then
            Debug.Print fso.GetBaseName(fichier.Name)
            nomSansExtension = fso.GetBaseName(fichier.Name)
            If InStr(facturesManquantesConnues, nomSansExtension & ".") = 0 Then
                dictPDFs(nomSansExtension) = False
            End If
        End If
    Next fichier

    Set ListerPDFs = dictPDFs
    
    'Libérer la mémoire
    Set dictPDFs = Nothing
    Set dossierObj = Nothing
    Set fichier = Nothing
    Set fso = Nothing
    
End Function

Sub zz_Comparer2Classeurs()
    
    Application.ScreenUpdating = False
    
    'Declare and open the 2 workbooks
    Dim wbWas As Workbook
    Set wbWas = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx", ReadOnly:=True)
    Debug.Print "#066 - " & wbWas.Name
    Dim wbNow As Workbook
    Set wbNow = Workbooks.Open("C:\VBA\GC_FISCALITÉ\GCF_DataFiles\2024_09_01_1835\GCF_BD_Entrée_TBA.xlsx", ReadOnly:=True)
    Debug.Print "#067 - " & wbNow.Name

    'Declare the 2 worksheets
    Dim wsWas As Worksheet
    Set wsWas = wbWas.Worksheets("Clients")
    Dim wsNow As Worksheet
    Set wsNow = wbNow.Worksheets("Clients")
    
    'Détermine la dernière ligne utilisée dans chacune des 2 feuilles
    Dim lastUsedRowWas As Long
    lastUsedRowWas = wsWas.Cells(wsWas.Rows.count, 1).End(xlUp).Row
    Dim lastUsedRowNOw As Long
    lastUsedRowNOw = wsNow.Cells(wsNow.Rows.count, 1).End(xlUp).Row
    
    'Détermine le nombre de colonnes dans l'ancienne feuille
    Dim lastUsedColWas As Long
    lastUsedColWas = wsWas.Cells(wsWas.Columns.count).End(xlToLeft).Column
    
    'Erase and create a new worksheet for differences
    Dim wsNameStr As String
    wsNameStr = "X_Différences"
    Dim wsDiff As Worksheet
    Call CreerOuRemplacerFeuille(wsNameStr)
    Set wsDiff = ThisWorkbook.Worksheets(wsNameStr)
    wsDiff.Range("A1").Value = "Ligne"
    wsDiff.Range("B1").Value = "Colonne"
    wsDiff.Range("C1").Value = "CodeClient"
    wsDiff.Range("D1").Value = "Nom du Client"
    wsDiff.Range("E1").Value = "Avant changement"
    wsDiff.Range("F1").Value = "Type"
    wsDiff.Range("G1").Value = "Après changement"
    Call CreerEnteteDeFeuille(wsDiff.Range("A1:G1"), RGB(0, 112, 192))

    Dim diffRow As Long
    diffRow = 2 'Take into consideration the Header
    Dim diffCol As Long
    diffCol = 1

    'Parcourir chaque ligne de l'ancienne version
    Dim cellWas As Range, cellNow As Range
    Dim foundRow As Range
    Dim clientCode As String
    Dim readCells As Long
    Dim i As Long, j As Long
    For i = 1 To lastUsedRowWas
        clientCode = CStr(wsWas.Cells(i, 2).Value)
        'Trouver la ligne correspondante dans la nouvelle version
        Set foundRow = wsNow.Columns(2).Find(What:=clientCode, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundRow Is Nothing Then
            Debug.Print "#068 - Ligne : " & i
            'Comparer les cellules des lignes correspondantes
            For j = 1 To lastUsedColWas
                readCells = readCells + 1
                Set cellWas = wsWas.Cells(i, j)
                Set cellNow = wsNow.Cells(foundRow.row, j)
                If CStr(cellWas.Value) <> CStr(cellNow.Value) Then
                    wsDiff.Cells(diffRow, 1).Value = i
                    wsDiff.Cells(diffRow, 2).Value = j
                    wsDiff.Cells(diffRow, 3).Value = wsWas.Cells(i, 2).Value
                    wsDiff.Cells(diffRow, 4).Value = wsWas.Cells(cellWas.row, 1).Value
                    wsDiff.Cells(diffRow, 5).Value = cellWas.Value
                    wsDiff.Cells(diffRow, 6).Value = "'--->"
                    wsDiff.Cells(diffRow, 7).Value = cellNow.Value
                    diffRow = diffRow + 1
                End If
            Next j
        Else
            wsDiff.Cells(diffRow, 1).Value = i
            wsDiff.Cells(diffRow, 3).Value = wsWas.Cells(i, 2).Value
            wsDiff.Cells(diffRow, 4).Value = wsWas.Cells(cellWas.row, 1).Value
            wsDiff.Cells(diffRow, 5).Value = cellWas.Value
            wsDiff.Cells(diffRow, 6).Value = "XXXX"
            diffRow = diffRow + 1
        End If
    Next i
            
    wsDiff.Columns.AutoFit
    
    'Result print setup - 2024-08-05 @ 05:16
    diffRow = diffRow + 1
    wsDiff.Range("A" & diffRow).Value = "**** " & Format$(readCells, "###,##0") & _
                                        " cellules analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:G" & diffRow)
    Call modAppli_Utils.AppliquerConditionalFormating(rngArea, 1, RGB(173, 216, 230))

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsDiff.Range("A2:DC" & diffRow)
    Dim header1 As String: header1 = "Vérification des différences"
    Dim header2 As String: header2 = "Clients"
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsDiff, rngToPrint, header1, header2, "$1:$1", "P")
    
    Application.ScreenUpdating = True
    
    wsDiff.Activate

    'Close the workbooks without saving
    wbWas.Close SaveChanges:=False
    wbNow.Close SaveChanges:=False
    
    'Libérer la mémoire
    Set cellWas = Nothing
    Set cellNow = Nothing
    Set foundRow = Nothing
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set wbWas = Nothing
    Set wbNow = Nothing
    Set wsWas = Nothing
    Set wsNow = Nothing
    Set wsDiff = Nothing
    
    MsgBox "La comparaison est complétée.", vbInformation
           
End Sub

Sub zz_DetecterErreurCodeClientInTEC()  '2025-03-11 @ 08:29

    'Source - Définir les chemins d'accès des fichiers, le Workbook et le Worksheet
    Dim sourceFilePath As String
    sourceFilePath = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                     wsdADMIN.Range("MASTER_FILE").Value
    Dim wbSource As Workbook: Set wbSource = Workbooks.Open(sourceFilePath)
    Dim wsSource As Worksheet: Set wsSource = wbSource.Worksheets("TEC_Local")
    
    'Détermine la dernière rangée et dernière colonne utilisées dans wsdTEC_Local
    Dim lastUsedRowTEC As Long
    lastUsedRowTEC = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).Row
    
    'Open the Master File Workbook
    Dim clientMFPath As String
    clientMFPath = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                     wsdADMIN.Range("CLIENTS_FILE").Value
    Dim wbMF As Workbook: Set wbMF = Workbooks.Open(clientMFPath)
    Dim wsMF As Worksheet: Set wsMF = wbMF.Worksheets("Clients")
    Dim lastUsedRowClient As Long
    lastUsedRowClient = wsMF.Cells(wsMF.Rows.count, 1).End(xlUp).Row
    
    'Setup output file
    Dim strOutput As String
    strOutput = "X_Détection_Cas_Erreur_Code_TEC"
    Call CreerOuRemplacerFeuille(strOutput)
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets(strOutput)
    wsOutput.Range("A1").Value = "TEC_ID"
    wsOutput.Range("B1").Value = "Date"
    wsOutput.Range("C1").Value = "Prof"
    wsOutput.Range("D1").Value = "NomClientTEC"
    wsOutput.Range("E1").Value = "CodeClient"
    wsOutput.Range("F1").Value = "NomClientFM"
    wsOutput.Range("G1").Value = "DateSaisie"
    Call CreerEnteteDeFeuille(wsOutput.Range("A1:G1"), RGB(0, 112, 192))
    
    'Build the dictionnary (Code, Nom du client) from Client's Master File
    Dim dictClients As Dictionary
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowClient
        dictClients.Add CStr(wsMF.Cells(i, fClntFMClientID).Value), wsMF.Cells(i, fClntFMClientNom).Value
    Next i
    
    'Parse TEC_Local to verify TEC's clientName vs. MasterFile's clientName
    Dim codeClientTEC As String, nomClientTEC As String, nomClientFromMF As String
    Dim casDelta As Long, rowOutput As Long
    rowOutput = 2
    For i = 2 To lastUsedRowTEC
        codeClientTEC = wsSource.Cells(i, fTECClientID).Value
        nomClientTEC = wsSource.Cells(i, fTECTDBClientNom).Value
        nomClientFromMF = dictClients(codeClientTEC)
        If Trim$(nomClientTEC) <> Trim$(nomClientFromMF) Then
            Debug.Print "#073 - " & i & " : " & codeClientTEC & " - " & nomClientTEC & " <---> " & nomClientFromMF
'            wsSource.Cells(i, 6).Value = nomClientFromMF
            wsOutput.Cells(rowOutput, 1).Value = wsSource.Cells(i, fTECTECID).Value
            wsOutput.Cells(rowOutput, 2).Value = wsSource.Cells(i, fTECDate).Value
            wsOutput.Cells(rowOutput, 3).Value = wsSource.Cells(i, fTECProf).Value
            wsOutput.Cells(rowOutput, 4).Value = nomClientTEC
            wsOutput.Cells(rowOutput, 5).Value = codeClientTEC
            wsOutput.Cells(rowOutput, 6).Value = nomClientFromMF
            wsOutput.Cells(rowOutput, 7).Value = wsSource.Cells(i, fTECDateSaisie).Value
            rowOutput = rowOutput + 1
            casDelta = casDelta + 1
        End If
    Next i
    
    wsOutput.Columns.AutoFit

    'Result print setup
    rowOutput = rowOutput + 1
    wsOutput.Range("A" & rowOutput).Value = "**** " & Format$(lastUsedRowTEC - 1, "###,##0") & _
                                        " lignes analysées dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsOutput.Range("A2:G" & rowOutput)
    Call modAppli_Utils.AppliquerConditionalFormating(rngArea, 1, RGB(173, 216, 230))

    'Setup print parameters
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:G" & rowOutput)
    Dim header1 As String: header1 = "Détection des codes de clients ERRONÉS dans TEC"
    Dim header2 As String: header2 = vbNullString
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    'Close the 2 workbooks without saving anything
    wbSource.Close SaveChanges:=False
    wbMF.Close SaveChanges:=False

    'Libérer la mémoire
    Set dictClients = Nothing
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set wbMF = Nothing
    Set wbSource = Nothing
    Set wsMF = Nothing
    Set wsOutput = Nothing
    Set wsSource = Nothing
    
    MsgBox _
        Prompt:="Il y a " & casDelta & " cas où le nom du client (TEC) diffère" & _
            vbNewLine & vbNewLine & "du nom de client du Fichier MAÎTRE", _
        Title:="Les données ne sont pas corrigées", _
        Buttons:=vbInformation
    
End Sub

Sub ReplacerFormesDepuisIntact()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim shpSrc As Shape, shpDest As Shape
    
    Set wsSource = ThisWorkbook.Worksheets("FAC_Finale_Intact")
    Set wsDest = ThisWorkbook.Worksheets("FAC_Finale")
    
    On Error Resume Next ' si une forme n'existe pas dans FAC_Finale
    For Each shpSrc In wsSource.Shapes
        Set shpDest = wsDest.Shapes(shpSrc.Name)
        If Not shpDest Is Nothing Then
            ' Repositionner et redimensionner
            shpDest.Top = shpSrc.Top
            shpDest.Left = shpSrc.Left
            shpDest.Width = shpSrc.Width
            shpDest.Height = shpSrc.Height
        End If
        Set shpDest = Nothing
    Next shpSrc
    On Error GoTo 0
    
    MsgBox "Formes replacées selon FAC_Finale_Intact.", vbInformation
End Sub


