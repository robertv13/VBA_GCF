Attribute VB_Name = "modDevTools"
Option Explicit

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

Sub Build_File_Layouts() '2024-03-26 @ 14:35

    Dim arr(1 To 20, 1 To 2) As Variant
    Dim output(1 To 150, 1 To 5) As Variant
    Dim r As Long
    r = 0
    r = r + 1: arr(r, 1) = "AR_Entête": arr(r, 2) = "A2:J2"
    r = r + 1: arr(r, 1) = "BD_Clients": arr(r, 2) = "A1:Q1"
    r = r + 1: arr(r, 1) = "Doc_ConditionalFormatting": arr(r, 2) = "A1:E1"
    r = r + 1: arr(r, 1) = "Doc_Formules": arr(r, 2) = "A1:H1"
    r = r + 1: arr(r, 1) = "Doc_NamedRanges": arr(r, 2) = "A1:B1"
    r = r + 1: arr(r, 1) = "Doc_Subs&Functions": arr(r, 2) = "A1:G1"
    r = r + 1: arr(r, 1) = "ENC_Entête": arr(r, 2) = "A3:F3"
    r = r + 1: arr(r, 1) = "ENC_Détail": arr(r, 2) = "A3:F3"
    r = r + 1: arr(r, 1) = "FAC_Entête": arr(r, 2) = "A3:T3"
    r = r + 1: arr(r, 1) = "FAC_Détails": arr(r, 2) = "A3:G3"
    r = r + 1: arr(r, 1) = "GL_Trans": arr(r, 2) = "A1:J1"
    r = r + 1: arr(r, 1) = "GL_EJ_Auto": arr(r, 2) = "C1:J1"
    r = r + 1: arr(r, 1) = "Invoice List": arr(r, 2) = "A2:J2"
    r = r + 1: arr(r, 1) = "TEC_Local": arr(r, 2) = "A2:P2"
    r = 1
    Dim i As Long, colNo As Long
    For i = 1 To UBound(arr, 1)
        If arr(i, 1) = "" Then Exit For
        Dim rng As Range: Set rng = Sheets(arr(i, 1)).Range(arr(i, 2))
        colNo = 0
        Dim cell As Range
        For Each cell In rng
            colNo = colNo + 1
            output(r, 2) = arr(i, 1)
            output(r, 3) = Chr(64 + colNo)
            output(r, 4) = colNo
            output(r, 5) = cell.value
            r = r + 1
        Next cell
    Next i
    
    'Setup and prepare the output worksheet
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Sheets("Doc_TableLayouts")
    Dim lastUsedRow As Long
    lastUsedRow = wsOutput.Range("A999").End(xlUp).row 'Last Used Row
    wsOutput.Range("A2:F" & lastUsedRow + 1).ClearContents
    
    wsOutput.Range("A2").Resize(r, 5).value = output
    
    'Libérer la mémoire
    Set rng = Nothing
    Set cell = Nothing
    Set wsOutput = Nothing
    
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
        wsName = wso.Name
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
            If col1.Font.Name <> col2.Font.Name Then
                diffLog = diffLog & "Column " & i & " Font Name differs: " & col1.Font.Name & " vs " & col2.Font.Name & vbCrLf
                wsDiff.Cells(diffRow, 3).value = i
                wsDiff.Cells(diffRow, 4).value = col1.Font.Name
                wsDiff.Cells(diffRow, 5).value = col2.Font.Name
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
    wsDiff.Cells(diffRow, 1).value = "Prod: " & wb1.Name
    diffRow = diffRow + 1
    wsDiff.Cells(diffRow, 1).value = "Dev : " & wb2.Name
    
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
            arr(nbColProd) = wsProd.Cells(1, nbColProd).value
            Debug.Print wsProd.Name, " Prod: ", wsProd.Cells(1, nbColProd).value
        Loop Until wsProd.Cells(1, nbColProd).value = ""
        nbColProd = nbColProd - 1
        nbRowProd = wsProd.Cells(wsProd.rows.count, "A").End(xlUp).row
        
        'Determine number of columns and rows in Dev Workbook
        Dim nbColDev As Integer, nbRowDev As Long
        nbColDev = 0
        Do
            nbColDev = nbColDev + 1
            Debug.Print wsDev.Name, " Dev : ", wsDev.Cells(1, nbColDev).value
        Loop Until wsProd.Cells(1, nbColDev).value = ""
        nbColDev = nbColDev - 1
        nbRowDev = wsDev.Cells(wsDev.rows.count, "A").End(xlUp).row
        
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

Sub LireFichierLogSaisieHeuresTXT() '2024-10-17 @ 20:13
    
    'Initialisation de la boîte de dialogue FileDialog pour choisir le fichier
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Configuration des filtres de fichiers (TXT uniquement)
    fd.Title = "Sélectionnez un fichier TXT"
    fd.Filters.Clear
    fd.Filters.Add "Fichiers Texte", "*.txt"
    
    'Si l'utilisateur sélectionne un fichier, filePath contiendra son chemin
    Dim filePath As String
    If fd.show = -1 Then
        filePath = fd.selectedItems(1)
    Else
        MsgBox "Aucun fichier sélectionné.", vbExclamation
        Exit Sub
    End If
    
    'Ouvre le fichier en mode lecture
    Dim FileNum As Integer
    FileNum = FreeFile
    Open filePath For Input As FileNum
    
    'Initialise la ligne de départ pour insérer les données dans Excel
    Dim ligneNum As Long
    ligneNum = 1
    
    'Lire chaque ligne du fichier
    Dim ligne As String
    Dim champs() As String
    Dim j As Long

    Do While Not EOF(FileNum)
        Line Input #FileNum, ligne
        
        'Séparer les champs par le séparateur " | "
        champs = Split(ligne, " | ")
        
        'Insérer les champs dans les colonnes de la feuille Excel
        For j = LBound(champs) To UBound(champs)
            Cells(ligneNum, j + 1).value = champs(j)
        Next j
        
        'Passer à la ligne suivante
        ligneNum = ligneNum + 1
    Loop
    
    'Fermer le fichier
    Close FileNum
    
    'Libérer la mémoire
    Set fd = Nothing
    
    MsgBox "Le fichier a été importé avec succès.", vbInformation
    
End Sub

Sub Fix_Date_Format()
    
    'Initialisation de la boîte de dialogue FileDialog pour choisir le fichier Excel
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Configuration des filtres de fichiers (Excel uniquement)
    fd.Title = "Sélectionnez un fichier Excel"
    fd.Filters.Clear
    fd.Filters.Add "Fichiers Excel", "*.xlsx; *.xlsm"
    
    'Si l'utilisateur sélectionne un fichier, filePath contiendra son chemin
    Dim filePath As String
    Dim fileSelected As Boolean
    If fd.show = -1 Then
        filePath = fd.selectedItems(1)
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
        
        'Ajouter des feuilles et colonnes spécifiques (exemple)
'        colonnesANettoyer.Add "DEB_Recurrent", Array("B") 'Vérifier la colonne B
'        colonnesANettoyer.Add "DEB_Trans", Array("B") 'Vérifier la colonne B
'
'        colonnesANettoyer.Add "ENC_Détails", Array("D") 'Vérifier la colonne D
'        colonnesANettoyer.Add "ENC_Entête", Array("B") 'Vérifier la colonne B
'
'        colonnesANettoyer.Add "FAC_Comptes_Clients", Array("B", "G") 'Vérifier et corriger les colonnes B & G
'        colonnesANettoyer.Add "FAC_Entête", Array("B") 'Vérifier et corriger la colonne B
'        colonnesANettoyer.Add "FAC_Projets_Détails", Array("F") 'Vérifier et corriger la colonne F
'        colonnesANettoyer.Add "FAC_Projets_Entête", Array("D") 'Vérifier et corriger la colonne D
'
'        colonnesANettoyer.Add "GL_Trans", Array("B") 'Vérifier et corriger la colonne B
'
'        colonnesANettoyer.Add "TEC_Local", Array("D") 'Vérifier et corriger la colonne D
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
            Debug.Print wsName
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                'Récupérer les colonnes à traiter pour cette feuille
                cols = colonnesANettoyer(wsName)
                
                'Parcourir chaque colonne spécifiée
                For Each col In cols
                    'Parcourir chaque cellule de la colonne spécifiée
                    For Each cell In ws.columns(col).SpecialCells(xlCellTypeConstants)
                        'Vérifier si la cellule contient une date avec une heure
                        If IsDate(cell.value) Then
                            'Vérifier si la valeur contient des heures (fraction décimale)
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

    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    Dim lurTEC As Long
    lurTEC = wsTEC.Cells(wsTEC.rows.count, "A").End(xlUp).row
    
    Dim wsTDB As Worksheet: Set wsTDB = wshTEC_TDB_Data
    Dim lurTDB As Long
    lurTDB = wsTDB.Cells(wsTDB.rows.count, "A").End(xlUp).row
    
    Dim wsOutput As Worksheet: Set wsOutput = wshzDocAnalyseÉcartTEC
    Dim lastUsed As Long
    lastUsed = wsOutput.Cells(wsOutput.rows.count, "A").End(xlUp).row + 2
    wsOutput.Range("A2:D" & lastUsed).ClearContents
    
    wsOutput.Cells(1, 1).value = "TECID"
    wsOutput.Cells(1, 2).value = "TEC_Local"
    wsOutput.Cells(1, 3).value = "TEC_TDB_Data"
    wsOutput.Cells(1, 4).value = "Vérification"
    
    Dim arr() As Variant
    ReDim arr(1 To 5000, 1 To 3)
    
    Dim i As Long
    Dim TECID As Long
    Dim dateCutOff As Date
    dateCutOff = Now()
    
    Dim h As Currency, hTEC As Currency
    'Boucle dans TEC_Local
    Debug.Print "Mise en mémoire TEC_LOCAL"
    For i = 3 To lurTEC
        With wsTEC
            If .Range("D" & i).value > dateCutOff Then Stop
            TECID = CLng(.Range("A" & i).value)
            If arr(TECID, 1) <> "" Then Stop
            arr(TECID, 1) = TECID
            h = .Range("H" & i).value
            If UCase(.Range("N" & i).value) = "VRAI" Then
                h = 0
            End If
            If h <> 0 Then
                If UCase(.Range("J" & i).value) = "VRAI" And Len(.Range("E" & i).value) > 2 Then
                    If UCase(.Range("L" & i).value) = "FAUX" Then
                        If .Range("M" & i).value <= dateCutOff Then
                            arr(TECID, 2) = h
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
    Debug.Print "Mise en mémoire TEC_TDB"
    For i = 2 To lurTDB
        With wsTDB
            If .Range("D" & i).value > dateCutOff Then Stop
            TECID = CLng(.Range("A" & i).value)
            arr(TECID, 1) = TECID
            arr(TECID, 3) = .Range("Q" & i).value
        End With
    Next i
    
    Debug.Print "Analyse des écarts"
    Dim tTEC As Double, tTDB As Double
    Dim r As Long: r = 2
    wsOutput.columns(2).EntireColumn.NumberFormat = "##0.00"
    wsOutput.Range("B:B").HorizontalAlignment = xlRight
    wsOutput.columns(3).EntireColumn.NumberFormat = "##0.00"
    wsOutput.Range("C:C").HorizontalAlignment = xlRight
    
    For i = 1 To 5000
        tTEC = tTEC + arr(i, 2)
        tTDB = tTDB + arr(i, 3)
        If arr(i, 2) <> 0 Or arr(i, 3) <> 0 Then
            wsOutput.Cells(r, 1).value = arr(i, 1)
            wsOutput.Cells(r, 2).value = arr(i, 2)
            wsOutput.Cells(r, 3).value = arr(i, 3)
            If arr(i, 2) <> arr(i, 3) Then
                wsOutput.Cells(r, 4).value = "Valeurs sont différentes"
            End If
            r = r + 1
        End If
    Next i
    
    wsOutput.Cells(r + 1, 2).value = Round(tTEC, 2)
    wsOutput.Cells(r + 1, 3).value = Round(tTDB, 2)
    
    Debug.Print "Totaux", Round(tTEC, 2), Round(tTDB, 2)
    
End Sub



