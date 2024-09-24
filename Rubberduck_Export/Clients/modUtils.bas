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
                Debug.Print "DB.#155 - workbook="; fullFileName; "  feuille="; feuilleName; "   colonne="; colName
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

Sub Test_Valider_Client_Avant_Effacement()

    Dim clientExiste As Boolean
    
    Call Valider_Client_Avant_Effacement("2026", clientExiste)
    
    Debug.Print clientExiste
    
End Sub
