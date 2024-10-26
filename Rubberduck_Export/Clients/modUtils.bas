Attribute VB_Name = "modUtils"
Option Explicit

Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Sub CM_Log_Record(moduleProcName As String, param1 As String, Optional ByVal startTime As Double = 0) '2024-08-22 @ 05:48

    Dim currentTime As String
    currentTime = Format$(Now, "yyyymmdd_hhmmss")
    
    'Determine the location of the Log file
    Dim rootPath As String
    If Fn_Get_Windows_Username <> "Robert M. Vigneault" Then
        rootPath = "P:\Administration\APP\GCF"
    Else
        rootPath = "C:\VBA\GC_FISCALITÉ"
    End If

    Dim logFile As String
    logFile = rootPath & DATA_PATH & Application.PathSeparator & "LogClientsApp.txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As #fileNum
    
    Dim moduleName As String, procName As String
    If InStr(moduleProcName, ":") Then
        moduleName = Left(moduleProcName, InStr(moduleProcName, ":") - 1)
        procName = Right(moduleProcName, Len(moduleProcName) - InStr(moduleProcName, ":"))
    Else
        moduleName = moduleProcName
        procName = ""
    End If
    
    If startTime = 0 Then
        startTime = Timer 'Start timing
        Print #fileNum, currentTime & "|" & _
                        ThisWorkbook.Name & "|" & _
                        Replace(Fn_Get_Windows_Username, " ", "_") & "|" & _
                        moduleName & "|" & _
                        procName & "|" & _
                        "" & "|" & _
                        param1
                        
    ElseIf startTime <= 0 Then 'Log intermédiaire
        Print #fileNum, currentTime & "|" & _
                        ThisWorkbook.Name & "|" & _
                        Replace(Fn_Get_Windows_Username, " ", "_") & "|" & _
                        moduleName & "|" & _
                        procName & "|" & _
                        "checkPoint" & "|" & _
                        param1
    Else
        Dim elapsedTime As Double
        elapsedTime = Round(Timer - startTime, 4) 'Calculate elapsed time
        Print #fileNum, currentTime & "|" & _
                        ThisWorkbook.Name & "|" & _
                        Replace(Fn_Get_Windows_Username, " ", "_") & "|" & _
                        moduleName & "|" & _
                        procName & " (sortie)" & "|" & _
                        "Temps écoulé: " & Format(elapsedTime, "#0.0000") & " secondes" & "|" & _
                        param1
    End If
    
    Close #fileNum

End Sub

Sub CM_Get_Date_Derniere_Modification(fileName As String, ByRef ddm As Date, _
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

Sub CM_Verify_DDM(fullFileName As String)

    Dim ddm As Date, jours As Long, heures As Long, minutes As Long, secondes As Long
    
    Call CM_Get_Date_Derniere_Modification(fullFileName, ddm, jours, heures, minutes, secondes)
    
    'Record to the log the difference between NOW and the date of last modifcation
    Call CM_Log_Record("modMain:CM_Update_External_GCF_BD_Entrée", "DDM (" & jours & "." & heures & "." & minutes & "." & secondes & ")", -1)
    If jours > 0 Or heures > 0 Or minutes > 0 Or secondes > 3 Then
        MsgBox "ATTENTION, le fichier MAÎTRE (GCF_Entrée.xlsx)" & vbNewLine & vbNewLine & _
               "n'a pas été modifié adéquatement sur disque..." & vbNewLine & vbNewLine & _
               "VEUILLEZ CONTACTER LE DÉVELOPPEUR SVP" & vbNewLine & vbNewLine & _
               "Code: (" & jours & "." & heures & "." & minutes & "." & secondes & ")", vbCritical, _
               "Le fichier n'est pas à jour sur disque"
    End If

End Sub

Sub Max_Code_Values_From_GCF_Entree(ByRef maxSmallCodes As String, ByRef maxLargeCodes As String)

    'Analyze Clients List from 'GCF_BD_Entrée.xlsx
    Dim strFilePath As String, strSheet As String
    If Not Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        strFilePath = "P:\Administration\APP\GCF\DataFiles\GCF_BD_Entrée.xlsx"
    Else
        strFilePath = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx"
    End If
    strSheet = "Clients$" 'Ne pas oublier le '$' à la fin du nom de la feuille
    
    'Crée une connexion à ADO
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    
    'Connexion pour Excel
    Dim strConn As String: strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                     "Data Source=" & strFilePath & ";" & _
                                     "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    'Ouvrir la connexion
    cn.Open strConn

    'Requête pour trouver la valeur maximale pour les codes de 1 à 999
    Dim sqlQuery As String
    sqlQuery = "SELECT MAX(Val(Client_ID)) AS MaxSmallCodes FROM [" & strSheet & "] WHERE Val(Client_ID) >= 1 AND Val(Client_ID) <= 999"
    Dim rs As Object
    Set rs = cn.Execute(sqlQuery)

    If Not rs.EOF Then
        maxSmallCodes = rs.Fields("MaxSmallCodes").Value
    Else
        maxSmallCodes = ""
    End If
    
    rs.Close

    'Requête pour trouver la valeur maximale pour les codes supérieurs ou égaux à 1000
    sqlQuery = "SELECT MAX(Val(Client_ID)) AS MaxLargeCodes FROM [" & strSheet & "] WHERE Len(Client_ID) >= 4 AND Val(Client_ID) >= 1000 AND Val(Client_ID) < 2000"
    Set rs = cn.Execute(sqlQuery)

    If Not rs.EOF Then
        maxLargeCodes = rs.Fields("MaxLargeCodes").Value
    Else
        maxLargeCodes = ""
    End If

    'Fermer le Recordset et la connexion
    rs.Close
    cn.Close
    
    If maxSmallCodes <> "" Then
        maxSmallCodes = Fn_Incremente_Code(maxSmallCodes)
    End If

    If maxLargeCodes <> "" Then
        maxLargeCodes = Fn_Incremente_Code(maxLargeCodes)
    End If

'    'Afficher les résultats
'    MsgBox "Valeur maximale pour les codes de 1 à 999: " & maxSmallCodes
'    MsgBox "Valeur maximale pour les codes >= 1000: " & maxLargeCodes
'
    'Nettoyer les objets
    Set rs = Nothing
    Set cn = Nothing
    
End Sub

Sub Valider_Client_Avant_Effacement(clientID As String, Optional ByRef clientExiste As Boolean = False) '2024-08-30 @ 18:15
    
    'Liste des workbooks à vérifier (à adapter selon vos besoins)
    Dim listeWorkbooks As Variant
    listeWorkbooks = Array("GCF_BD_MASTER.xlsx")
    
    Dim dataFilesPath As String
    If Not Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        dataFilesPath = "P:\Administration\APP\GCF\DataFiles"
    Else
        dataFilesPath = "C:\VBA\GC_FISCALITÉ\DataFiles"
    End If

    'Boucle pour vérifier dans les workbooks fermés
    Dim fullFileName As String, message1 As String, message2 As String
    Dim sql As String
    Dim conn As Object
    Dim rs As Object
    Dim i As Integer
    For i = LBound(listeWorkbooks) To UBound(listeWorkbooks)
        fullFileName = dataFilesPath & "\" & listeWorkbooks(i)
        
        'Vérifier l'existence du fichier
        If Dir(fullFileName) <> "" Then
            'Utiliser ADO pour ouvrir le workbook fermé
            Set conn = CreateObject("ADODB.Connection")
            conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
            
            'Boucle sur les feuilles à vérifier (exemple: "Sheet1", "Sheet2")
            Dim feuilleRechercher As Variant
            Dim plageRechercher As String, colName As String, feuilleName As String
            For Each feuilleRechercher In Array("ENC_Entête|codeClient", _
                                                "FAC_Comptes_Clients|CodeClient", _
                                                "FAC_Entête|Cust_ID", _
                                                "FAC_Projets_Détails|ClientID", _
                                                "FAC_Projets_Entête|ClientID", _
                                                "TEC_Local|Client_ID")
                colName = Mid(feuilleRechercher, InStr(feuilleRechercher, "|") + 1)
                feuilleName = Left(feuilleRechercher, InStr(feuilleRechercher, "|") - 1)
                plageRechercher = feuilleName & "$"
                
                ' Construire la requête SQL pour chercher le client
                sql = "SELECT * FROM [" & plageRechercher & "] WHERE [" & colName & "] = '" & clientID & "'"
                
                Set rs = conn.Execute(sql)
                If Not rs.EOF Then
                    message1 = message1 & "Le client '" & clientID & "' existe dans la feuille '" & feuilleName & "'" & vbCrLf
                    clientExiste = True
                GoTo Exit_Sub
                End If
                rs.Close
            Next feuilleRechercher
            
            conn.Close
        End If
    Next i
    
    'Boucle pour vérifier dans les worksheets du workbook actif
    Dim wb As Workbook
    
    For Each wb In Application.Workbooks
        If wb.Name = "Vérification de la liste de clients.xlsx" Then
            GoTo Next_Workbook
        End If
        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            Dim foundCell As Range
            If ws.Name = "Données" Or ws.Name = "DonnéesRecherche" Or ws.Name = "Clients" Then
                GoTo Next_Worksheet
            End If
            Set foundCell = ws.Cells.Find(What:=clientID, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                message2 = message2 & "Le client '" & clientID & "' existe dans la feuille '" & ws.Name & "' du Workbook '" & wb.Name & "'" & vbCrLf
                clientExiste = True
                GoTo Exit_Sub
            End If
Next_Worksheet:
        Next ws
Next_Workbook:
    Next wb
    
    'clean up
    Set conn = Nothing
    Set foundCell = Nothing
    Set rs = Nothing
    Set wb = Nothing
    Set ws = Nothing

Exit_Sub:
    If message1 <> "" Then
        MsgBox message1, vbCritical, "Ce code de client est utilisé dans le fichier MASTER"
    End If
    If message2 <> "" Then
        MsgBox message2, vbCritical, "Ce code de client est utilisé dans le fichier Clients"
    End If
    
End Sub

Sub Code_Search_Everywhere() '2024-10-26 @ 11:27
    
    'Declare lineOfCode() as variant
    Dim allLinesOfCode As Variant
    ReDim allLinesOfCode(1 To 25000, 1 To 4)
    
'    Application.ScreenUpdating = False
    
    'Allows up to 3 search strings
    Dim search1 As String, search2 As String, search3 As String
    search1 = InputBox("Enter the search string ? ", "Search1")
    search2 = InputBox("Enter the search string ? ", "Search2")
    search3 = InputBox("Enter the search string ? ", "Search3")
    
'    Application.ScreenUpdating = True
    
    'Loop through all VBcomponents (modules, class and forms) in the active workbook
    Dim lineNum As Long
    Dim X As Long
    
    Dim vbComp As Object
    Dim oType As String
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
        Case 1
            oType = "1_Module"
        Case 2
            oType = "2_Class"
        Case 3
            oType = "3_userform"
        Case 100
            oType = "0_Worksheet"
        Case Else
            oType = oType & "_?????"
            Stop
        End Select
        
        'Get the code module for the component
        Dim vbCodeMod As Object: Set vbCodeMod = vbComp.CodeModule
        
        'Loop through all lines in the code module to save all the lines in memory
        For lineNum = 1 To vbCodeMod.CountOfLines
            If Trim(vbCodeMod.Lines(lineNum, 1)) <> "" Then
                X = X + 1
                allLinesOfCode(X, 1) = oType
                allLinesOfCode(X, 2) = vbComp.Name
                allLinesOfCode(X, 3) = lineNum
                allLinesOfCode(X, 4) = Trim(vbCodeMod.Lines(lineNum, 1))
            End If
        Next lineNum
    Next vbComp
    
    'At this point allLinesOfCode contains all lines of code of the application - 2024-07-10 @ 17:33
    
    Call Array_2D_Resizer(allLinesOfCode, X, UBound(allLinesOfCode, 2))
    
    Call Search_Every_Lines_Of_Code(allLinesOfCode, search1, search2, search3)
    
    'Clean up
    Set vbComp = Nothing
    Set vbCodeMod = Nothing
    
End Sub

Sub Array_2D_Resizer(ByRef inputArray As Variant, ByVal nRows As Long, ByVal nCols As Long)
    
    Dim oRows As Long, oCols As Long
    
    'Get the original dimensions of the input array
    oRows = UBound(inputArray, 1)
    oCols = UBound(inputArray, 2)
    
    'Ensure the new dimensions are within the original array's bounds
    If nRows > oRows Then nRows = oRows
    If nCols > oCols Then nCols = oCols
    
    'Create a new array with the specified dimensions
    Dim tempArray() As Variant
    ReDim tempArray(1 To nRows, 1 To nCols)
    
    ' Copy the relevant data from the input array to the new array
    Dim i As Long, j As Long
    For i = 1 To nRows
        For j = 1 To nCols
            tempArray(i, j) = inputArray(i, j)
        Next j
    Next i
    
    ' Assign the trimmed array back to the input array
    inputArray = tempArray
    
End Sub

Sub Search_Every_Lines_Of_Code(arr As Variant, search1 As String, search2 As String, search3 As String)

    'Declare arr() to keep results in memory
    Dim arrResult() As Variant
    ReDim arrResult(1 To 2000, 1 To 7)

    Dim posProcedure As Long, posFunction As Long
    Dim saveLineOfCode As String, trimmedLineOfCode As String, procedureName As String
    Dim TimeStamp As String
    Dim X As Long, xr As Long
    For X = LBound(arr, 1) To UBound(arr, 1)
        trimmedLineOfCode = arr(X, 4)
        saveLineOfCode = trimmedLineOfCode
        
        'Handle comments (second parameter is either Remove or Uppercase)
        If InStr(1, trimmedLineOfCode, "'") <> 0 Then
            trimmedLineOfCode = HandleComments(trimmedLineOfCode, "U")
        End If
        
        If trimmedLineOfCode <> "" Then
            'Is this a procedure (Sub) declaration line ?
            If InStr(trimmedLineOfCode, "Sub ") <> 0 Then
                If InStr(trimmedLineOfCode, "End Sub") = 0 And _
                    InStr(trimmedLineOfCode, "Sub = ") = 0 And _
                    InStr(trimmedLineOfCode, "Sub As ") = 0 And _
                    InStr(trimmedLineOfCode, "Exit Sub") = 0 Then
                        procedureName = Mid(saveLineOfCode, InStr(trimmedLineOfCode, "Sub "))
                End If
            End If
            
            If InStr(trimmedLineOfCode, "End Sub") = 1 Then
                procedureName = ""
            End If

            'Is this a function declaration line ?
            If InStr(trimmedLineOfCode, "Function ") <> 0 Then
                If InStr(trimmedLineOfCode, "End Function") = 0 And _
                    InStr(trimmedLineOfCode, "Function = ") = 0 And _
                    InStr(trimmedLineOfCode, "Function As ") = 0 And _
                    InStr(trimmedLineOfCode, "Exit Function") = 0 Then
                        procedureName = Mid(saveLineOfCode, InStr(trimmedLineOfCode, "Function "))
                End If
            End If
            
            If InStr(trimmedLineOfCode, "End Function") = 1 Then
                procedureName = ""
            End If
            
            'Do we find the search1 or search2 or sreach3 strings in this line of code ?
            If (search1 <> "" And InStr(trimmedLineOfCode, search1) <> 0) Or _
                (search2 <> "" And InStr(trimmedLineOfCode, search2) <> 0) Or _
                (search3 <> "" And InStr(trimmedLineOfCode, search3) <> 0) Then
                'Found an occurence
                xr = xr + 1
                arrResult(xr, 2) = arr(X, 1) 'oType
                arrResult(xr, 3) = arr(X, 2) 'oName
                arrResult(xr, 4) = arr(X, 3) 'LineNum
                arrResult(xr, 5) = procedureName
                arrResult(xr, 6) = "'" & saveLineOfCode
                TimeStamp = Format$(Now(), "mm/dd/yyyy hh:mm:ss")
                arrResult(xr, 7) = TimeStamp
                arrResult(xr, 1) = UCase(arr(X, 1)) & Chr(0) & UCase(arr(X, 2)) & Chr(0) & Format$(arr(X, 3), "0000") & Chr(0) & procedureName 'Future sort key
            End If
        End If
    Next X

    'Prepare the result worksheet
    Call Erase_And_Create_Worksheet("X_Doc_Search_Utility_Results")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Doc_Search_Utility_Results")
    wsOutput.Range("A1").Value = "SortKey"
    wsOutput.Range("B1").Value = "Type"
    wsOutput.Range("C1").Value = "ModuleName"
    wsOutput.Range("D1").Value = "LineNo"
    wsOutput.Range("E1").Value = "ProcedureName"
    wsOutput.Range("F1").Value = "Code"
    wsOutput.Range("G1").Value = "TimeStamp"
    
    Call Make_It_As_Header(wsOutput.Range("A1:G1"))
    
    'Is there anything to show ?
    If xr > 0 Then
    
        'Data starts at row 2
        Dim r As Long: r = 2

        Call Array_2D_Resizer(arrResult, xr, UBound(arrResult, 2))
        
        'Sort the 2D array based on column 1
        Call Array_2D_Bubble_Sort(arrResult)
    
        'Transfer the array to the worksheet
        wsOutput.Range("A2").Resize(UBound(arrResult, 1), UBound(arrResult, 2)).Value = arrResult
        wsOutput.Range("A:A").EntireColumn.Hidden = True 'Do not show the sortKey
        wsOutput.Columns(4).HorizontalAlignment = xlCenter
        wsOutput.Columns(7).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        Dim lastUsedRow As Long
        lastUsedRow = wsOutput.Range("B9999").End(xlUp).Row
        Dim j As Long, oldProcedure As String
        oldProcedure = wsOutput.Range("C" & lastUsedRow).Value & wsOutput.Range("E" & lastUsedRow).Value
        For j = lastUsedRow To 2 Step -1
            If wsOutput.Range("C" & j).Value & wsOutput.Range("E" & j).Value <> oldProcedure Then
                wsOutput.Rows(j + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                oldProcedure = wsOutput.Range("C" & j).Value & wsOutput.Range("E" & j).Value
            End If
        Next j
        
        'Since we might have inserted new row, let's update the lastUsedRow
        lastUsedRow = wsOutput.Range("B9999").End(xlUp).Row
        With wsOutput.Range("B2:G" & lastUsedRow)
            On Error Resume Next
            Cells.FormatConditions.Delete
            On Error GoTo 0
        
            .FormatConditions.Add Type:=xlExpression, Formula1:= _
                "=(MOD(LIGNE();2)=1)"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
            End With
            .FormatConditions(1).StopIfTrue = False
        End With
        
        wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit
    End If
    
    'Result print setup - 2024-07-14 2 06:24
    lastUsedRow = lastUsedRow + 2
    wsOutput.Range("B" & lastUsedRow).Value = "*** " & Format$(X, "###,##0") & " lignes de code dans l'application ***"
    Dim header1 As String: header1 = "Search Utility Results"
    Dim header2 As String
    header2 = "Searched strings '" & search1 & "'"
    If search2 <> "" Then header2 = header2 & " '" & search2 & "'"
    If search3 <> "" Then header2 = header2 & " '" & search3 & "'"
    Call Simple_Print_Setup(wsOutput, wsOutput.Range("B2:G" & lastUsedRow), _
                           header1, _
                           header2, _
                           "$1:$1", _
                           "L")
    
    'Display the final message
    If xr Then
        MsgBox "J'ai trouvé " & xr & " lignes avec les chaines '" & search1 & "'" & vbNewLine & _
                vbNewLine & "après avoir analysé un total de " & _
                Format$(X, "#,##0") & " lignes de code"
    Else
        MsgBox "Je n'ai trouvé aucune occurences avec les chaines '" & search1 & "'" & vbNewLine & _
                vbNewLine & "après avoir analysé un total de " & _
                Format$(X, "#,##0") & " lignes de code"
    End If
    
    'Libérer la mémoire
    Set wsOutput = Nothing
    
End Sub

Function HandleComments(ByVal codeLine As String, action As String) As String '2024-06-30 @ 10:45
    
    'R as action will remove the comments
    'U as action will UPPERCASE the comments
    
    Dim inString As Boolean: inString = False
    Dim codePart As String, commentPart As String
    
    Debug.Assert action = "R" Or action = "U"
    
    Dim i As Long, char As String
    For i = 1 To Len(codeLine)
        char = Mid(codeLine, i, 1)
        
        'Toggle inString flag if a double quote is encountered
        If char = """" Then
            inString = Not inString
        End If
        
        'If the current character is ' and we are not within a string...
        If char = "'" Then
            If Not inString Then
                commentPart = Mid(codeLine, i)
                Exit For
            Else
                codePart = codePart & char
            End If
        Else
            codePart = codePart & char
        End If
    Next i
    
    'Take action - R remove the comment from the code, L uppercase the comment
    If action = "R" Then
        commentPart = ""
    Else
        commentPart = Trim(UCase(commentPart))
    End If
    
    HandleComments = codePart & commentPart
    
End Function

Sub Erase_And_Create_Worksheet(sheetName As String)

    Dim ws As Worksheet
    Dim wsExists As Boolean

    'Check if the worksheet exists
    wsExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
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

    'Create a new worksheet with the specified name
    Set ws = ThisWorkbook.Worksheets.Add(Before:=wshMENU)
    ws.Name = sheetName
    
    'Libérer la mémoire
    Set ws = Nothing
    
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
    wsName = r.Worksheet.Name
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(wsName)
    ws.Columns.AutoFit
    
    'Libérer la mémoire
    Set r = Nothing
    Set ws = Nothing

End Sub

Sub Array_2D_Bubble_Sort(ByRef arr() As Variant) '2024-06-23 @ 07:05
    
    Dim i As Long, j As Long, numRows As Long, numCols As Long
    Dim Temp As Variant
    Dim sorted As Boolean
    
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    'Bubble Sort Algorithm
    Dim c As Long, cProcess As Long
    For i = 1 To numRows - 1
        sorted = True
        For j = 1 To numRows - i
            'Compare column 2 first
            If arr(j, 1) > arr(j + 1, 1) Then
                'Swap rows
                For c = 1 To numCols
                    Temp = arr(j, c)
                    arr(j, c) = arr(j + 1, c)
                    arr(j + 1, c) = Temp
                Next c
                sorted = False
            ElseIf arr(j, 1) = arr(j + 1, 1) Then
                'Column 1 values are equal, then compare column2 values
                If arr(j, 2) > arr(j + 1, 2) Then
                    'Swap rows
                    For c = 1 To numCols
                        Temp = arr(j, c)
                        arr(j, c) = arr(j + 1, c)
                        arr(j + 1, c) = Temp
                    Next c
                    sorted = False
                End If
            End If
        Next j
        'If no swaps were made, the array is sorted
        If sorted Then Exit For
    Next i

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



