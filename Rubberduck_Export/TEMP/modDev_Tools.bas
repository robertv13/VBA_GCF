Attribute VB_Name = "modDev_Tools"
Option Explicit

Sub Get_Range_From_Dynamic_Named_Range(dynamicRangeName As String, ByRef rng As Range)
    
    On Error Resume Next
    'R�cup�rer la formule associ�e au nom
    Dim refersToFormula As String
    refersToFormula = ThisWorkbook.Names(dynamicRangeName).RefersTo
    On Error GoTo 0
    
    If refersToFormula = "" Then
        MsgBox "La plage nomm�e '" & dynamicRangeName & "' n'existe pas ou est invalide.", vbExclamation
        Exit Sub
    End If
    
    'Tester et �valuer la plage
    On Error Resume Next
    Set rng = Application.Evaluate(refersToFormula)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Impossible de r�soudre la plage nomm�e dynamique '" & dynamicRangeName & "'. V�rifiez la d�finition.", vbExclamation
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
        MsgBox "Il existe des r�f�rences circulaires dans le Workbook dans les cellules suivantes:" & vbCrLf & circRef, vbExclamation
    Else
        MsgBox "Il n'existe aucune r�f�rence circulaire dans ce Workbook .", vbInformation
    End If
    
    'Lib�rer la m�moire
    Set cell = Nothing
    Set formulaCells = Nothing
    Set ws = Nothing
    
End Sub

Sub Build_File_Layouts() '2024-03-26 @ 14:35

    Dim arr(1 To 20, 1 To 2) As Variant
    Dim output(1 To 150, 1 To 5) As Variant
    Dim r As Long
    r = 0
    r = r + 1: arr(r, 1) = "AR_Ent�te": arr(r, 2) = "A2:J2"
    r = r + 1: arr(r, 1) = "BD_Clients": arr(r, 2) = "A1:Q1"
    r = r + 1: arr(r, 1) = "Doc_ConditionalFormatting": arr(r, 2) = "A1:E1"
    r = r + 1: arr(r, 1) = "Doc_Formules": arr(r, 2) = "A1:H1"
    r = r + 1: arr(r, 1) = "Doc_NamedRanges": arr(r, 2) = "A1:B1"
    r = r + 1: arr(r, 1) = "Doc_Subs&Functions": arr(r, 2) = "A1:G1"
    r = r + 1: arr(r, 1) = "ENC_Ent�te": arr(r, 2) = "A3:F3"
    r = r + 1: arr(r, 1) = "ENC_D�tail": arr(r, 2) = "A3:F3"
    r = r + 1: arr(r, 1) = "FAC_Ent�te": arr(r, 2) = "A3:T3"
    r = r + 1: arr(r, 1) = "FAC_D�tails": arr(r, 2) = "A3:G3"
    r = r + 1: arr(r, 1) = "GL_Trans": arr(r, 2) = "A1:J1"
    r = r + 1: arr(r, 1) = "GL_EJ_R�currente": arr(r, 2) = "C1:J1"
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
            output(r, 5) = cell.Value
            r = r + 1
        Next cell
    Next i
    
    'Setup and prepare the output worksheet
    Dim wsOutput As Worksheet: Set wsOutput = wshzDocTableLayouts
    Dim lastUsedRow As Long
    lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "A").End(xlUp).row 'Last Used Row
    wsOutput.Range("A2:F" & lastUsedRow + 1).ClearContents
    
    wsOutput.Range("A2").Resize(r, 5).Value = output
    
    'Lib�rer la m�moire
    Set rng = Nothing
    Set cell = Nothing
    Set wsOutput = Nothing
    
End Sub

Sub Compare_2_Workbooks_Column_Formatting()                      '2024-08-19 @ 16:24

    'Erase and create a new worksheet for differences
    Dim wsDiff As Worksheet
    Call CreateOrReplaceWorksheet("Diff�rences_Colonnes")
    Set wsDiff = ThisWorkbook.Worksheets("Diff�rences_Colonnes")
    wsDiff.Range("A1").Value = "Worksheet"
    wsDiff.Range("B1").Value = "Nb. colonnes"
    wsDiff.Range("C1").Value = "Colonne"
    wsDiff.Range("D1").Value = "Valeur originale"
    wsDiff.Range("E1").Value = "Nouvelle valeur"
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
            If col1.Interior.color <> col2.Interior.color Then
                diffLog = diffLog & "Column " & i & " Background Color differs: " & col1.Interior.color & " vs " & col2.Interior.color & vbCrLf
                wsDiff.Cells(diffRow, 3).Value = i
                wsDiff.Cells(diffRow, 4).Value = col1.Interior.color
                wsDiff.Cells(diffRow, 5).Value = col2.Interior.color
            End If
    
        Next i
        
    Next wso
    
    wsDiff.Columns.AutoFit
    wsDiff.Range("B:E").Columns.HorizontalAlignment = xlCenter
    
    'Result print setup - 2024-08-05 @ 05:16
    diffRow = diffRow + 2
    wsDiff.Range("A" & diffRow).Value = "**** " & Format$(readColumns, "###,##0") & _
                                        " colonnes analys�es dans l'ensemble du fichier ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:E" & diffRow)
    Call modAppli_Utils.ApplyConditionalFormatting(rngArea, 1, True)

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
        MsgBox "Diff�rences trouv�es:" & vbCrLf & diffLog
    Else
        MsgBox "Aucune diff�rence dans les colonnes."
    End If
    
    'Lib�rer la m�moire
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
    wsDiff.Range("A1").Value = "Worksheet"
    wsDiff.Range("B1").Value = "Prod_Cols"
    wsDiff.Range("C1").Value = "Dev_Cols"
    wsDiff.Range("D1").Value = "Prod_Rows"
    wsDiff.Range("E1").Value = "Dev_Rows"
    wsDiff.Range("F1").Value = "Ligne #"
    wsDiff.Range("G1").Value = "Colonne"
    wsDiff.Range("H1").Value = "Prod_Value"
    wsDiff.Range("I1").Value = "Dev_Value"
    Call Make_It_As_Header(wsDiff.Range("A1:I1"))

    'Set your workbooks and worksheets here
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open("C:\VBA\GC_FISCALIT�\GCF_DataFiles\GCF_BD_MASTER_COPY.xlsx")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER.xlsx")
    
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
        nbRowProd = wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row
        
        'Determine number of columns and rows in Dev Workbook
        Dim nbColDev As Integer, nbRowDev As Long
        nbColDev = 0
        Do
            nbColDev = nbColDev + 1
            Debug.Print "#045 - " & wsDev.Name, " Dev : ", wsDev.Cells(1, nbColDev).Value
        Loop Until wsProd.Cells(1, nbColDev).Value = ""
        nbColDev = nbColDev - 1
        nbRowDev = wsDev.Cells(wsDev.Rows.count, 1).End(xlUp).row
        
        diffRow = diffRow + 2
        wsDiff.Cells(diffRow, 1).Value = wsName
        wsDiff.Cells(diffRow, 2).Value = nbColProd
        wsDiff.Cells(diffRow, 3).Value = nbColDev
        wsDiff.Cells(diffRow, 4).Value = nbRowProd
        wsDiff.Cells(diffRow, 5).Value = nbRowDev
        
        Dim nbRow As Long
        If nbRowProd > nbRowDev Then
            wsDiff.Cells(diffRow, 6).Value = "Le client a ajout� " & nbRowProd - nbRowDev & " lignes dans la feuille"
            nbRow = nbRowProd
        End If
        If nbRowProd < nbRowDev Then
            wsDiff.Cells(diffRow, 6).Value = "Le dev a ajout� " & nbRowDev - nbRowProd & " lignes dans la feuille"
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
                                        " lignes analys�es dans l'ensemble du Workbook ***"
                                    
    'Set conditional formatting for the worksheet (alternate colors)
    Dim rngArea As Range: Set rngArea = wsDiff.Range("A2:I" & diffRow)
    Call modAppli_Utils.ApplyConditionalFormatting(rngArea, 1, True)

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
        MsgBox "Diff�rences trouv�es:" & vbCrLf & diffLogMess
    Else
        MsgBox "Aucune diff�rence dans les lignes."
    End If
    
    'Lib�rer la m�moire
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
    
    'Initialisation de la bo�te de dialogue FileDialog pour choisir le fichier
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Configuration des filtres de fichiers (TXT uniquement)
    fd.Title = "S�lectionnez un fichier TXT"
    fd.Filters.Clear
    fd.Filters.Add "Fichiers Texte", "*.txt"
    
    'Si l'utilisateur s�lectionne un fichier, filePath contiendra son chemin
    Dim FilePath As String
    If fd.show = -1 Then
        FilePath = fd.selectedItems(1)
    Else
        MsgBox "Aucun fichier s�lectionn�.", vbExclamation
        Exit Sub
    End If
    
    'Ouvre le fichier en mode lecture
    Dim FileNum As Integer
    FileNum = FreeFile
    Open FilePath For Input As FileNum
    
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
            Cells(ligneNum, j + 1).Value = champs(j)
        Next j
        
        'Passer � la ligne suivante
        ligneNum = ligneNum + 1
    Loop
    
    'Fermer le fichier
    Close FileNum
    
    'Lib�rer la m�moire
    Set fd = Nothing
    
    MsgBox "Le fichier a �t� import� avec succ�s.", vbInformation
    
End Sub

Sub Fix_Date_Format()
    
    'Initialisation de la bo�te de dialogue FileDialog pour choisir le fichier Excel
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Configuration des filtres de fichiers (Excel uniquement)
    fd.Title = "S�lectionnez un fichier Excel"
    fd.Filters.Clear
    fd.Filters.Add "Fichiers Excel", "*.xlsx; *.xlsm"
    
    'Si l'utilisateur s�lectionne un fichier, filePath contiendra son chemin
    Dim FilePath As String
    Dim fileSelected As Boolean
    If fd.show = -1 Then
        FilePath = fd.selectedItems(1)
        fileSelected = True
    Else
        MsgBox "Aucun fichier s�lectionn�.", vbExclamation
        fileSelected = False
    End If
    
    'Ouvrir le fichier s�lectionn� s'il y en a un
    Dim wb As Workbook
    If fileSelected Then
        Set wb = Workbooks.Open(FilePath)
        
        'D�finir les colonnes sp�cifiques � nettoyer pour chaque feuille
        Dim colonnesANettoyer As Dictionary
        Set colonnesANettoyer = CreateObject("Scripting.Dictionary")
        
        'Ajouter des feuilles et colonnes sp�cifiques (exemple)
'        colonnesANettoyer.add "DEB_Recurrent", Array("B") 'V�rifier la colonne B
'        colonnesANettoyer.add "DEB_Trans", Array("B") 'V�rifier la colonne B
'
'        colonnesANettoyer.add "ENC_D�tails", Array("D") 'V�rifier la colonne D
'        colonnesANettoyer.add "ENC_Ent�te", Array("B") 'V�rifier la colonne B
'
'        colonnesANettoyer.add "FAC_Comptes_Clients", Array("B", "G") 'V�rifier et corriger les colonnes B & G
'        colonnesANettoyer.add "FAC_Ent�te", Array("B") 'V�rifier et corriger la colonne B
'        colonnesANettoyer.add "FAC_Projets_D�tails", Array("F") 'V�rifier et corriger la colonne F
'        colonnesANettoyer.add "FAC_Projets_Ent�te", Array("D") 'V�rifier et corriger la colonne D
'
'        colonnesANettoyer.add "GL_Trans", Array("B") 'V�rifier et corriger la colonne B
'
'        colonnesANettoyer.add "TEC_Local", Array("D") 'V�rifier et corriger la colonne D
        colonnesANettoyer.Add "TEC_Local", Array("M") 'V�rifier et corriger la colonne D
        
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
            Debug.Print "#046 - " & wsName
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                'R�cup�rer les colonnes � traiter pour cette feuille
                cols = colonnesANettoyer(wsName)
                
                'Parcourir chaque colonne sp�cifi�e
                For Each col In cols
                    'Parcourir chaque cellule de la colonne sp�cifi�e
                    For Each cell In ws.Columns(col).SpecialCells(xlCellTypeConstants)
                        'V�rifier si la cellule contient une date avec une heure
                        If IsDate(cell.Value) Then
                            'V�rifier si la valeur contient des heures (fraction d�cimale)
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
    
    'Lib�rer la m�moire
    Set cell = Nothing
    Set col = Nothing
    Set colonnesANettoyer = Nothing
    Set fd = Nothing
    Set wb = Nothing
    Set ws = Nothing
    Set wsName = Nothing
    
    MsgBox "Les dates ont �t� corrig�es pour les colonnes sp�cifiques.", vbInformation

End Sub

Sub Debug_�cart_TEC_Local_vs_TEC_TDB_Data()

    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    Dim lurTEC As Long
    lurTEC = wsTEC.Cells(wsTEC.Rows.count, 1).End(xlUp).row
    
    Dim wsTDB As Worksheet: Set wsTDB = wshTEC_TDB_Data
    Dim lurTDB As Long
    lurTDB = wsTDB.Cells(wsTDB.Rows.count, 1).End(xlUp).row
    
    Dim wsOutput As Worksheet: Set wsOutput = wshzDocAnalyse�cartTEC
    Dim lastUsed As Long
    lastUsed = wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).row + 2
    wsOutput.Range("A2:D" & lastUsed).ClearContents
    
    wsOutput.Cells(1, 1).Value = "TECID"
    wsOutput.Cells(1, 2).Value = "TEC_Local"
    wsOutput.Cells(1, 3).Value = "TEC_TDB_Data"
    wsOutput.Cells(1, 4).Value = "V�rification"
    
    Dim arr() As Variant
    ReDim arr(1 To 5000, 1 To 3)
    
    Dim i As Long
    Dim TECID As Long
    Dim dateCutOff As Date
    dateCutOff = Now()
    
    Dim h As Currency, hTEC As Currency
    'Boucle dans TEC_Local
    Debug.Print "#048 - Mise en m�moire TEC_LOCAL"
    For i = 3 To lurTEC
        With wsTEC
            If .Range("D" & i).Value > dateCutOff Then Stop
            TECID = CLng(.Range("A" & i).Value)
            If arr(TECID, 1) <> "" Then Stop
            arr(TECID, 1) = TECID
            h = .Range("H" & i).Value
            If UCase(.Range("N" & i).Value) = "VRAI" Then
                h = 0
            End If
            If h <> 0 Then
                If UCase(.Range("J" & i).Value) = "VRAI" And Len(.Range("E" & i).Value) > 2 Then
                    If UCase(.Range("L" & i).Value) = "FAUX" Then
                        If .Range("M" & i).Value <= dateCutOff Then
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
    Debug.Print "#049 - Mise en m�moire TEC_TDB"
    For i = 2 To lurTDB
        With wsTDB
            If .Range("D" & i).Value > dateCutOff Then Stop
            TECID = CLng(.Range("A" & i).Value)
            arr(TECID, 1) = TECID
            arr(TECID, 3) = .Range("Q" & i).Value
        End With
    Next i
    
    Debug.Print "#050 - Analyse des �carts"
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
                wsOutput.Cells(r, 4).Value = "Valeurs sont diff�rentes"
            End If
            r = r + 1
        End If
    Next i
    
    wsOutput.Cells(r + 1, 2).Value = Round(tTEC, 2)
    wsOutput.Cells(r + 1, 3).Value = Round(tTDB, 2)
    
    'Lib�rer la m�moire
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
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    r = 2
    
    Dim ligneCode As String, moduleName As String, procName As String
    Dim ObjetSet As String, objetForEach As String, objetNothing As String
    
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
        ligneCode = Trim(ws.Cells(i, 6))
'        If InStr(ligneCode, "= Nothing") Then
'            If InStr(ligneCode, " recSet ") = 0 Then
'                ligneCode = Replace(ligneCode, "Set", "set")
'            End If
'        End If
        If InStr(ligneCode, "recSet As ") Then
            ligneCode = Replace(ligneCode, "recSet As ", "resste As ")
        End If
        If InStr(ligneCode, ".Recordset") Then
            ligneCode = Replace(ligneCode, ".Recordset", ".RecordSET")
        End If
        If InStr(ligneCode, ".offset") Then
            ligneCode = Replace(ligneCode, ".offset", ".offSET")
        End If
        If InStr(ligneCode, ".Offset") Then
            ligneCode = Replace(ligneCode, ".Offset", ".OffSET")
        End If
        
        ObjetSet = ""
        objetForEach = ""
        objetNothing = ""
        'D�claration de l'objet avec Set...
        If InStr(ligneCode, "Set ") <> 0 Then
            If Left(ligneCode, 4) = "Set " Or InStr(ligneCode, ": Set") <> 0 Then
                ObjetSet = Mid(ligneCode, InStr(ligneCode, "Set ") + 4, Len(ligneCode))
                ObjetSet = Left(ObjetSet, InStr(ObjetSet, " ") - 1)
                If ObjetSet = "As" Then Stop
                If InStr(added, ObjetSet & "|") = 0 Then
                    added = added + ObjetSet + "|"
                End If
            Else
                Debug.Print ligneCode
            End If
        End If
        'D�claration de l'objet avec For Each...
        If InStr(ligneCode, "For Each ") <> 0 Then
            objetForEach = Mid(ligneCode, InStr(ligneCode, "For Each ") + 9, Len(ligneCode))
            objetForEach = Left(objetForEach, InStr(objetForEach, " ") - 1)
            If objetForEach = "As" Then Stop
            If InStr(added, objetForEach & "|") = 0 Then
                added = added + objetForEach + "|"
            End If
        End If
        'Lib�ration de l'objet avec = Nothing
        If InStr(ligneCode, " = Nothing") <> 0 Then
            objetNothing = Mid(ligneCode, InStr(ligneCode, "Set") + 4, Len(ligneCode))
            objetNothing = Left(objetNothing, InStr(objetNothing, " ") - 1)
            If objetNothing = "" Then Stop
            cleared = cleared + objetNothing + "|"
        End If
        
Next_For:
    Next i
    
    'Lib�rer la m�moire
    Set ws = Nothing
    Set wsOutput = Nothing
    
End Sub

Sub Sauvegarder_UserForms_Parameters() '2024-11-26 @ 07:42

    'Utiliser la feuille 'UserForm_Params' ou la cr�er pour sauvegarder les param�tres
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wshUserFormParams
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "UserForm_Params"
    End If
    On Error GoTo 0
    
    'En-t�tes de colonnes
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Nom_UserForm", "Largeur", "Hauteur", "Position_Left", "Position_Top")
    
    Dim i As Integer
    i = 2
    'Parcourir tous les composants VBA pour trouver les UserForms
    Dim VBComp As Object
    Dim userFormName As String
    Dim uf As Object
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        If VBComp.Type = vbext_ct_MSForm Then
            userFormName = VBComp.Name
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
                ' D�charger le UserForm pour lib�rer la m�moire
                Unload uf
                Set uf = Nothing
            End If
        End If
    Next VBComp
    
    'Lib�rer la m�moire
    Set uf = Nothing
    
    MsgBox "Param�tres des UserForms sauvegard�s avec succ�s.", vbInformation

End Sub

Sub Restaurer_UserForms_Parameters()

    Dim ws As Worksheet
    Dim i As Integer
    Dim uf As Object
    Dim nomUF As String

    'V�rifier si la feuille existe
    On Error Resume Next
    Set ws = wshUserFormParams
    If ws Is Nothing Then
        MsgBox "La feuille 'UserForm_Params' n'existe pas. Sauvegardez d'abord les param�tres.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    'Parcourir la liste des param�tres sauvegard�s
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
            'Optionnel : afficher le UserForm pour v�rifier
            'uf.Show
        End If

        i = i + 1
    Loop

    'Lib�rer la m�moire
    Set uf = Nothing
    Set ws = Nothing
    
    MsgBox "Param�tres des UserForms restaur�s avec succ�s.", vbInformation

End Sub

Sub Get_UsedRange_In_Active_Workbook()

    Dim output As String
    
    'Feuille pour les r�sultats
    Dim feuilleNom As String
    feuilleNom = "X_Cellules_Utilis�es"
    Call Erase_And_Create_Worksheet(feuilleNom)
    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Sheets(feuilleNom)
    Dim r As Long: r = 1
    wsOutput.Cells(r, 1).Value = "Feuille"
    wsOutput.Cells(r, 2).Value = "Plage utilis�e"
    wsOutput.Cells(r, 3).Value = "Lignes utilis�e"
    wsOutput.Cells(r, 4).Value = "Colonnes utilis�e"
    wsOutput.Cells(r, 5).Value = "Nb. Cellules"
    r = r + 1
    
    'Parcourir chaque feuille du classeur
    Dim ws As Worksheet
    Dim cellCount As Long
    For Each ws In ThisWorkbook.Worksheets
        'V�rifier si UsedRange n'est pas vide
        On Error Resume Next
        Dim usedRange As Range
        Set usedRange = ws.usedRange
        On Error GoTo 0
        
        If Not usedRange Is Nothing Then
            ' Ajouter les informations � la sortie
            wsOutput.Cells(r, 1).Value = ws.Name
            wsOutput.Cells(r, 2).Value = usedRange.Address
            wsOutput.Cells(r, 3).Value = usedRange.Rows.count
            wsOutput.Cells(r, 4).Value = usedRange.Columns.count
            wsOutput.Cells(r, 5).Value = usedRange.Cells.count
        Else
            ' Si aucune cellule utilis�e
            wsOutput.Cells(r, 1).Value = ws.Name
            wsOutput.Cells(r, 2).Value = "Aucune"
        End If
        r = r + 1
    Next ws
    
    MsgBox "Le traitement est compl�t�. Voir la feuille '" & feuilleNom & "'", vbInformation
    
End Sub

Sub CreerRepertoireEtImporterFichiers() '2024-12-09 @ 22:26

    'Chemin du dossier contenant les fichiers PROD
    Dim cheminSourcePROD As String
    cheminSourcePROD = "P:\Administration\APP\GCF\DataFiles\" ' Ajustez ce chemin
    
    'V�rifier si des fichiers Actif_*.txt existent (utilisateurs encore pr�sents)
    Dim actifFile As String
    Dim actifExists As Boolean
    actifFile = Dir(cheminSourcePROD & "Actif_*.txt")
    actifExists = (actifFile <> "")
    
    If actifExists Then
        MsgBox "Un ou plusieurs utilisateurs utilisent encore l'application." & vbNewLine & vbNewLine & _
               "La copie est annul�e.", vbExclamation
        Exit Sub
    End If
    
    'D�finir le chemin racine (local) pour la cr�ation du nouveau dossier
    Dim cheminRacineDestination As String
    cheminRacineDestination = "C:\VBA\GC_FISCALIT�\GCF_DataFiles\"
    
    'Construire le nom du r�pertoire bas� sur la date et l'heure actuelle
    Dim dateHeure As String
    Dim nouveauDossier As String
    dateHeure = Format(Now, "yyyy_mm_dd_hhnn")
    nouveauDossier = cheminRacineDestination & dateHeure & "\"
    
    'Cr�er le r�pertoire s'il n'existe pas d�j� (ne devrait pas exister)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.folderExists(nouveauDossier) Then
        FSO.CreateFolder nouveauDossier
    End If
    
    'Noms des deux fichiers � copier (fixe)
    Dim nomFichier1 As String, nomFichier2 As String
    nomFichier1 = "GCF_BD_MASTER.xlsx"
    nomFichier2 = "GCF_BD_Entr�e.xlsx"
    
    'Copier le premier fichier
    If FSO.fileExists(cheminSourcePROD & nomFichier1) Then
        FSO.CopyFile source:=cheminSourcePROD & nomFichier1, Destination:=nouveauDossier, OverwriteFiles:=False
    Else
        MsgBox "Fichier non trouv� : " & cheminSourcePROD & nomFichier1, vbExclamation, "Erreur"
    End If
    
    'Copier le deuxi�me fichier
    If FSO.fileExists(cheminSourcePROD & nomFichier2) Then
        FSO.CopyFile source:=cheminSourcePROD & nomFichier2, Destination:=nouveauDossier, OverwriteFiles:=False
    Else
        MsgBox "Fichier non trouv� : " & cheminSourcePROD & nomFichier2, vbExclamation, "Erreur"
    End If

    'Copier les fichiers .log (variable)
    Dim fichier As String
    fichier = Dir(cheminSourcePROD & "*.log")
    Do While fichier <> ""
        'Copie du fichier PROD ---> Local
        FSO.CopyFile source:=cheminSourcePROD & fichier, Destination:=nouveauDossier, OverwriteFiles:=False
        'Efface le fichier PROD (initialiation)
        Kill cheminSourcePROD & fichier
        'Fichier suivant � copier
        fichier = Dir
    Loop
    
    'Copie des deux fichiers du dossier temporaire vers le dossier DEV (but ultime)
    
    Dim dossierDEV As String
    dossierDEV = "C:\VBA\GC_FISCALIT�\DataFiles\"
    
    'Copier le premier fichier
    If FSO.fileExists(nouveauDossier & nomFichier1) Then
        FSO.CopyFile source:=cheminSourcePROD & nomFichier1, Destination:=dossierDEV, OverwriteFiles:=True
    Else
        MsgBox "Fichier non trouv� : " & nouveauDossier & nomFichier1, vbExclamation, "Erreur"
    End If
    
    'Copier le deuxi�me fichier
    If FSO.fileExists(nouveauDossier & nomFichier2) Then
        FSO.CopyFile source:=cheminSourcePROD & nomFichier2, Destination:=dossierDEV, OverwriteFiles:=True
    Else
        MsgBox "Fichier non trouv� : " & nouveauDossier & nomFichier2, vbExclamation, "Erreur"
    End If

    MsgBox "Fichiers copi�s dans le dossier : " & nouveauDossier, vbInformation, "Termin�"

End Sub

Sub AjusterTableauxDansMaster() '2024-12-07 @ 06:47

    'Chemin du classeur � ajuster
    Dim cheminClasseur As String
    cheminClasseur = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER.xlsx"

    If Fn_Get_Windows_Username <> "Robert M. Vigneault" Then
        Exit Sub
    End If
    
    'Ouvrir le classeur
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks.Open(cheminClasseur, ReadOnly:=False)
    If wb Is Nothing Then
        MsgBox "Impossible d'ouvrir le classeur 'GCF_BD_MASTER.xlsx'", vbExclamation, "Erreur"
        Exit Sub
    End If
    On Error GoTo 0

    'Parcourir toutes les feuilles
    Dim ws As Worksheet
    Dim listeObjets As ListObjects
    Dim tableau As ListObject
    Dim derniereLigne As Long
    Dim derniereColonne As Long
    Dim nouvellePlage As Range
    
    For Each ws In wb.Worksheets
        Set listeObjets = ws.ListObjects
        'Parcourir chaque tableau de la feuille
        For Each tableau In listeObjets
            'Trouver la derni�re ligne avec des donn�es
            derniereLigne = ws.Cells(ws.Rows.count, tableau.Range.Column).End(xlUp).row
            'Trouver la derni�re colonne avec des donn�es
            derniereColonne = ws.Cells(tableau.HeaderRowRange.row, ws.Columns.count).End(xlToLeft).Column
            'Red�finir la plage du tableau
            Set nouvellePlage = ws.Range(ws.Cells(tableau.HeaderRowRange.row, tableau.Range.Column), _
                                         ws.Cells(derniereLigne, derniereColonne))
            On Error Resume Next
            tableau.Resize nouvellePlage
            On Error GoTo 0
        Next tableau
    Next ws

    'Enregistrer et fermer le classeur
    wb.Save
    wb.Close
    
    MsgBox "Tous les tableaux ont �t� ajust�s avec succ�s.", vbInformation, "Termin�"
    
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
    Dim VBComp As Object
    Dim codeModule As Object
    Dim ligne As Long
    Dim found As Boolean
    Dim oleObj As OLEObject
    
    ' Parcourir toutes les feuilles du classeur
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print "V�rification des contr�les sur la feuille : " & ws.Name
        
        ' V�rification des Shapes (Formulaires ou Boutons assign�s)
        For Each shp In ws.Shapes
            On Error Resume Next
            macroNameRaw = shp.OnAction
            On Error GoTo 0
            
            If macroNameRaw <> "" Then
                ' Extraire uniquement le nom de la macro apr�s le "!"
                If InStr(1, macroNameRaw, "!") > 0 Then
                    macroName = Split(macroNameRaw, "!")(1)
                Else
                    macroName = macroNameRaw
                End If
                
                ' V�rifier si la macro existe
                found = VerifierMacroExiste(macroName)
                
                ' R�sultat de la v�rification
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
        
        ' V�rification des contr�les ActiveX
        For Each oleObj In ws.OLEObjects
            If TypeOf oleObj.Object Is MSForms.CommandButton Then
                ' Construire le nom de la macro � partir du nom du contr�le
                macroName = oleObj.Name & "_Click"
                
                ' V�rifier si la macro existe
                found = VerifierMacroExiste(macroName, ws.CodeName)
                
                ' R�sultat de la v�rification
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
    
    MsgBox "V�rification termin�e sur toutes les feuilles. Consultez la fen�tre Ex�cution pour les r�sultats.", vbInformation
    
End Sub

Function VerifierMacroExiste(macroName As String, Optional moduleName As String = "") As Boolean

    'Par defaut...
    VerifierMacroExiste = False
    
    'Si un module sp�cifique est fourni, v�rifier uniquement dans ce module
    Dim VBComp As Object
    Dim codeModule As Object
    Dim ligne As Long
    
    If moduleName <> "" Then
        On Error Resume Next
        Set VBComp = ThisWorkbook.VBProject.VBComponents(moduleName)
        On Error GoTo 0
        If Not VBComp Is Nothing Then
            Set codeModule = VBComp.codeModule
            For ligne = 1 To codeModule.CountOfLines
                If codeModule.ProcOfLine(ligne, vbext_pk_Proc) = macroName Then
                    VerifierMacroExiste = True
                    Exit Function
                End If
            Next ligne
        End If
        Exit Function
    End If
    
    'Parcourir tous les modules si aucun module sp�cifique n'est fourni
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        Set codeModule = VBComp.codeModule
        For ligne = 1 To codeModule.CountOfLines
            If codeModule.ProcOfLine(ligne, vbext_pk_Proc) = macroName Then
                VerifierMacroExiste = True
                Exit Function
            End If
        Next ligne
    Next VBComp
    
End Function

Sub Main() '2024-12-25 @ 15:27

    'Feuille pour la sortie
    Dim outputName As String
    outputName = "Doc_File_Layouts"
    Call CreateOrReplaceWorksheet(outputName)
    
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets(outputName)
    
    'Tableau pour travailler en m�moire les r�sultats
    Dim arrOut() As String
    ReDim arrOut(1 To 250, 1 To 7)
    
    Dim outputRow As Long
    outputRow = 1
    
    Call ListeEnumsGenerique("BD_Clients", 1, arrOut, outputRow)
    Call ListeEnumsGenerique("BD_Fournisseurs", 1, arrOut, outputRow)
    
    Call ListeEnumsGenerique("DEB_R�current", 1, arrOut, outputRow)
    Call ListeEnumsGenerique("DEB_Trans", 1, arrOut, outputRow)
    
    Call ListeEnumsGenerique("ENC_D�tails", 1, arrOut, outputRow)
    Call ListeEnumsGenerique("ENC_Ent�te", 1, arrOut, outputRow)
    
    Call ListeEnumsGenerique("FAC_Comptes_Clients", 2, arrOut, outputRow)
    Call ListeEnumsGenerique("FAC_D�tails", 2, arrOut, outputRow)
    Call ListeEnumsGenerique("FAC_Ent�te", 2, arrOut, outputRow)
    Call ListeEnumsGenerique("FAC_Projets_D�tails", 1, arrOut, outputRow)
    Call ListeEnumsGenerique("FAC_Projets_Ent�te", 1, arrOut, outputRow)
    Call ListeEnumsGenerique("FAC_Sommaire_Taux", 1, arrOut, outputRow)
    
    Call ListeEnumsGenerique("GL_EJ_R�currente", 1, arrOut, outputRow)
    Call ListeEnumsGenerique("GL_Trans", 1, arrOut, outputRow)
    
    Call ListeEnumsGenerique("TEC_Local", 2, arrOut, outputRow)
    Call ListeEnumsGenerique("TEC_TDB_Data", 1, arrOut, outputRow)
    
    '�criture des r�sultats (tableau) dans la feuille
    With wsOut
        .Cells.Clear 'Efface tout le contenu de la feuille
        .Range("A1").Resize(outputRow, UBound(arrOut, 2)).Value = arrOut
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
        Set wb = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx")
        tableName = Replace(tableName, "BD_", "")
    Else
        Set wb = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_MASTER.xlsx")
    End If
    Dim wsMaster As Worksheet
    If tableName <> "TEC_TDB_Data" Then
        Set wsMaster = wb.Sheets(tableName)
    End If
    tableName = saveTableName
    
    'Nom de la table
    arrArg(outputRow, 1) = tableName
    outputRow = outputRow + 1
    
    'Extraire la d�finition des Enum de la table � partir du code
    Dim arr() As Variant
    Call ExtractEnumDefinition(tableName, arr)
    
    'Boucle sur les colonnes
    Dim col As Long
    For col = LBound(arr, 1) To UBound(arr, 1)
        arrArg(outputRow, 1) = arr(col, 1)
        arrArg(outputRow, 2) = arr(col, 2)
        'Nom de la colonne dans la table
        arrArg(outputRow, 3) = ws.Cells(HeaderRow, col).Value
        If InStr(arr(col, 2), ws.Cells(HeaderRow, col).Value) = 0 Then
            arrArg(outputRow, 4) = "*"
        End If
        If Not wsMaster Is Nothing Then
            arrArg(outputRow, 5) = wsMaster.Cells(1, col).Value
            If InStr(arr(col, 2), wsMaster.Cells(1, col).Value) = 0 Then
                arrArg(outputRow, 6) = "*"
            End If
        End If
        'Valeurs des colonnes sur la premi�re ligne de data
        arrArg(outputRow, 7) = ws.Cells(HeaderRow + 1, col).Value
        outputRow = outputRow + 1
    Next col
    
    'Ligne pour s�parer les tables
    outputRow = outputRow + 1
    
    'Fermer sans sauvegarder
    wb.Close SaveChanges:=False
    
End Sub

Sub ExtractEnumDefinition(tableName As String, ByRef arr() As Variant)

    Dim LineNum As Long
    Dim TotalLines As Long
    Dim codeLine As String
    Dim InEnumBlock As Boolean
    Dim FilePath As String
    
    'Variable de travail
    Dim EnumDefinition As String
    EnumDefinition = ""
    
    'Redimensionner le tableau
    ReDim arr(1 To 50, 1 To 2)
    Dim e As Long
    
    'Acc�der au projet VBA actif
    Dim VBProj As VBIDE.VBProject
    Set VBProj = ThisWorkbook.VBProject

    'Parcourir tous les composants VBA
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.codeModule
    For Each VBComp In VBProj.VBComponents
        Set CodeMod = VBComp.codeModule
        'Parcourir chaque ligne de code
        For LineNum = 1 To CodeMod.CountOfLines
            codeLine = Trim(CodeMod.Lines(LineNum, 1))
            'D�tection du d�but d'un Enum
            If InStr(1, codeLine, "Enum " & tableName, vbTextCompare) > 0 Then
                InEnumBlock = True
            ElseIf InEnumBlock Then
                'D�tection de la fin de l'Enum
                If InStr(1, codeLine, "End Enum", vbTextCompare) > 0 Then
                    InEnumBlock = False
                    Exit For 'Terminer apr�s l'extraction
                Else
                    'Ajouter les lignes � l'int�rieur du Enum
                    If Left(codeLine, 1) <> "[" Then
                        If Right(codeLine, 11) = " = [_First]" Then
                            codeLine = Left(codeLine, Len(codeLine) - 11)
                        End If
                        e = e + 1
                        arr(e, 1) = e
                        arr(e, 2) = codeLine
                        EnumDefinition = EnumDefinition & codeLine & "|"
                    End If
                End If
            End If
        Next LineNum
    Next VBComp

    'Redimension au minimum le tableau
    Call Array_2D_Resizer(arr, e, 2)
    
End Sub
