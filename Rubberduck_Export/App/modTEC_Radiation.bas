Attribute VB_Name = "modTEC_Radiation"
Option Explicit

Public previousCellAddress As Variant

Sub TEC_Radiation_Procedure(codeClient As String, cutoffDate As String)

    If cutoffDate = "" Then
        Exit Sub
    End If
        
    Call TEC_Import_All
    
    Dim maxDate As Date
    Dim y As Integer, m As Integer, d As Integer
    y = year(cutoffDate)
    m = month(cutoffDate)
    d = day(cutoffDate)
    maxDate = DateSerial(y, m, d)
    
    Dim ws As Worksheet: Set ws = wshTEC_Radiation
    Dim wsSource As Worksheet: Set wsSource = wshTEC_Local
    
    'AdvancedFilter # 2 avec TEC_Local (Heures Facturables, Non Facturées, Non Détruites à la date Limite)
    Call Get_TEC_For_Client_AF(codeClient, CDate(cutoffDate), "VRAI", "FAUX", "FAUX")
    
    'Avons-nous des résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, "AQ").End(xlUp).row
    If lastUsedRow < 3 Then
        MsgBox "Il n'y a aucune TEC pour ce client", vbInformation
        Call Prepare_Pour_Nouvelle_Radiation
        wshTEC_Radiation.Range("F3").Activate
        GoTo ExitSub
    End If
    
    'Transfère la table en mémoire (arr)
    Dim arr As Variant
    arr = wsSource.Range("AQ3:BF" & lastUsedRow).value
    
    Dim i As Long
    Dim tecID As Long
    Dim dateTEC As Date
    Dim profInit As String, descTEC As String
    Dim hresTEC As Currency, tauxHoraire As Currency, valeurTEC As Currency
    Dim totalHresTEC As Currency, totalValeurTEC As Currency
    Dim currRow As Integer, activeRow As Integer
    currRow = 6
    For i = 1 To UBound(arr, 1)
        If currRow <= 30 Then
            tecID = CLng(arr(i, fTECTECID))
            dateTEC = Format$(arr(i, fTECDate), wshAdmin.Range("B1").value)
            profInit = Format$(arr(i, fTECProfID), "000") & arr(i, fTECProf)
            descTEC = arr(i, fTECDescription)
            hresTEC = arr(i, fTECHeures)
            With ws
                .Cells(currRow, 5).value = tecID
                .Cells(currRow, 6).value = dateTEC
                .Cells(currRow, 7).value = Mid(profInit, 4)
                .Cells(currRow, 8).value = descTEC
                .Cells(currRow, 10).value = hresTEC
                tauxHoraire = Fn_Get_Hourly_Rate(CLng(Left(profInit, 3)), CDate(cutoffDate))
                valeurTEC = hresTEC * tauxHoraire
                .Cells(currRow, 11).value = valeurTEC
            End With
            activeRow = currRow
            totalHresTEC = totalHresTEC + hresTEC
            totalValeurTEC = totalValeurTEC + valeurTEC
        End If
        currRow = currRow + 1
    Next i
    
    'La ligne maximum ne doit pas excéder 32
    currRow = currRow + 1
    If currRow > 32 Then
        currRow = 32
    End If
    
    'Affiche les totaux
    With ws
        .Cells(currRow, 8).value = "* TOTAUX *"
        .Cells(currRow, 8).Font.Bold = True
        .Cells(currRow, 10).Font.Bold = True
        .Cells(currRow, 11).Font.Bold = True
    End With
    
    Call AjouterCheckBoxesAvecControleGlobal(activeRow)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ws.Shapes("Impression").Visible = True
    ws.Shapes("Radiation").Visible = True
    
ExitSub:

    'Libérer la mémoire
    Set ws = Nothing
    Set wsSource = Nothing
    
End Sub

Sub AjouterCheckBoxesAvecControleGlobal(lastUsedRow)

    Dim ws As Worksheet
    Set ws = wshTEC_Radiation
    
    Dim i As Long
    Dim chkBox As checkBox
    Dim headerChkBox As checkBox

    'Supprimer les cases à cocher existantes (au cas où)
    For Each chkBox In ws.CheckBoxes
        chkBox.Delete
    Next chkBox

    'Ajouter une case à cocher dans l'en-tête pour tout cocher/décocher
    Set headerChkBox = ws.CheckBoxes.Add(Left:=ws.Cells(5, 4).Left + 5, _
                                         Top:=ws.Cells(5, 4).Top, _
                                         Width:=ws.Cells(5, 4).Width, _
                                         Height:=ws.Cells(5, 4).Height)
    With headerChkBox
        .Name = "chk_header"
        .Caption = ""
        .OnAction = "ToutCocherOuDecocher" 'Associe la macro de contrôle global
    End With

    'Ajouter une case à cocher pour chaque ligne du tableau
    For i = 6 To lastUsedRow
        Set chkBox = ws.CheckBoxes.Add(Left:=ws.Cells(i, 4).Left + 5, _
                                       Top:=ws.Cells(i, 4).Top, _
                                       Width:=ws.Cells(i, 4).Width, _
                                       Height:=ws.Cells(i, 4).Height)
        With chkBox
            .Name = "chk_" & i
            .Caption = ""
            .linkedCell = ws.Cells(i, 2).Address
            .OnAction = "CalculerTotaux"
        End With
    Next i

End Sub

Sub ToutCocherOuDecocher()

    Dim ws As Worksheet
    Set ws = wshTEC_Radiation ' Remplacez par le nom de votre feuille
    
    Dim headerChkBox As checkBox
    Dim chkBox As checkBox
    Dim newState As Boolean
    Dim allChecked As Boolean
    
    'Déprotéger la feuille temporairement
    ws.Unprotect
    
    'Trouver la case à cocher de l'en-tête
    On Error Resume Next
    Set headerChkBox = ws.CheckBoxes("chk_header")
    On Error GoTo 0
    'Vérifier si la case de l'en-tête existe
    If headerChkBox Is Nothing Then
        MsgBox "La case à cocher d'en-tête 'chk_header' n'existe pas.", vbExclamation
        Exit Sub
    End If
    
    'Vérifier si toutes les cases sous-jacentes sont cochées
    allChecked = True
    For Each chkBox In ws.CheckBoxes
        If chkBox.Name <> "chk_header" Then
            If chkBox.value <> xlOn Then
                allChecked = False
                Exit For
            End If
        End If
    Next chkBox
    
    'Déterminer le nouvel état à appliquer aux cases sous-jacentes
    newState = Not allChecked ' Si toutes sont cochées, on décoche tout, sinon on coche tout

    'Désactiver les événements
    Application.EnableEvents = False

    'Réinitialiser l'état de la case à cocher de l'en-tête pour être sûr qu'elle peut être modifiée
    headerChkBox.value = xlOff 'On s'assure que la case est décochée avant de basculer
    headerChkBox.value = IIf(newState, xlOn, xlOff) 'Appliquer l'état approprié
    
    'Parcourir toutes les cases à cocher et appliquer le nouvel état
    For Each chkBox In ws.CheckBoxes
        If chkBox.Name <> "chk_header" Then
            chkBox.value = IIf(newState, xlOn, xlOff)
            chkBox.Characters.Text = ""
            chkBox.Characters.Caption = ""
        End If
    Next chkBox

    'Réactiver les événements
    Application.EnableEvents = True

    'Protéger la feuille à nouveau
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                AllowInsertingColumns:=True, AllowInsertingRows:=True, AllowDeletingColumns:=True, _
                AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                AllowUsingPivotTables:=True
                
    Call CalculerTotaux
    
End Sub

Sub CalculerTotaux()

    Dim ws As Worksheet
    Set ws = wshTEC_Radiation
    
    Dim chkBox As checkBox
    Dim totalHres As Currency, totalValeur As Currency
    Dim rowNum As Long
    
    rowNum = 6
    For Each chkBox In ws.CheckBoxes
        If chkBox.Name <> "chk_header" Then
            'Vérifier si la case à cocher est cochée
            If chkBox.value = xlOn Then
                'Ajouter la valeur de la cellule correspondante à la somme totale
                totalHres = totalHres + ws.Cells(rowNum, "J").value
                totalValeur = totalValeur + ws.Cells(rowNum, "K").value
                
            End If
            rowNum = rowNum + 1
        End If
    Next chkBox
    
    ' Afficher le total dans une cellule spécifique (par exemple, cellule Z1)
    ws.Cells(rowNum + 2, 10).value = Format$(totalHres, "###,##0.00")
    ws.Cells(rowNum + 2, 11).value = Format$(totalValeur, "###,##0.00 $")
    
End Sub

Sub shp_TEC_Radiation_GO_Click()

    Call Radiation_Mise_À_Jour
    
    Call Prepare_Pour_Nouvelle_Radiation
    
    wshTEC_Radiation.Activate

End Sub

Sub Radiation_Mise_À_Jour()

    'Avons-nous des résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = wshTEC_Local.Cells(wshTEC_Local.Rows.count, "AQ").End(xlUp).row
    
    If lastUsedRow >= 3 Then
        Call TEC_Radiation_Update_As_Billed_To_DB(3, lastUsedRow)
        Call TEC_Radiation_Update_As_Billed_Locally(3, lastUsedRow)
    End If
    
    MsgBox "Les TEC sélectionnés ont été radié avec succès", vbOKOnly, "Confirmation de traitement"
    
End Sub

Sub TEC_Radiation_Update_As_Billed_To_DB(firstRow As Long, lastRow As Long) 'Update Billed Status in DB

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Radiation:TEC_Radiation_Update_As_Billed_To_DB", firstRow & ", " & lastRow, 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "TEC_Local$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim r As Long, tecID As Long, strSQL As String
    'Offset entre TEC_Local & wshTEC_Radiation
    Dim offset As Long
    offset = 3
    For r = firstRow To lastRow
        If wshTEC_Radiation.Range("B" & r + offset).value = True Then
            tecID = wshTEC_Local.Range("AQ" & r).value
            
            'Open the recordset for the specified ID
            strSQL = "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & tecID
            rs.Open strSQL, conn, 2, 3
            If Not rs.EOF Then
                'Update EstFacturee, DateFacturee & NoFacture
                rs.Fields(fTECEstFacturee - 1).value = "VRAI"
                rs.Fields(fTECDateFacturee - 1).value = Format$(Date, "yyyy-mm-dd")
                rs.Fields(fTECNoFacture - 1).value = "Radiation"
                rs.Update
            Else
                'Handle the case where the specified ID is not found
                MsgBox "L'enregistrement avec le TECID '" & r & "' ne peut être trouvé!", _
                    vbExclamation
                rs.Close
                conn.Close
                Exit Sub
            End If
            'Update the recordset (create the record)
            rs.Update
            rs.Close
        End If
next_iteration:
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modTEC_Radiation:TEC_Radiation_Update_As_Billed_To_DB", "", startTime)

End Sub

Sub TEC_Radiation_Update_As_Billed_Locally(firstResultRow As Long, lastResultRow As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Radiation:TEC_Radiation_Update_As_Billed_Locally", firstResultRow & ", " & lastResultRow, 0)
    
    'Set the range to look for
    Dim lookupRange As Range: Set lookupRange = wshTEC_Local.Range("l_tbl_TEC_Local[TECID]")
    
    Dim r As Long, rowToBeUpdated As Long, tecID As Long
    'Offset entre TEC_Local & wshTEC_Radiation
    Dim offset As Long
    offset = 3
    For r = firstResultRow To lastResultRow
        If wshTEC_Radiation.Range("B" & r + offset).value = True Then
            tecID = wshTEC_Local.Range("AQ" & r).value
            rowToBeUpdated = Fn_Find_Row_Number_TECID(tecID, lookupRange)
            wshTEC_Local.Cells(rowToBeUpdated, fTECEstFacturee).value = "VRAI"
            wshTEC_Local.Cells(rowToBeUpdated, fTECDateFacturee).value = Format$(Date, "yyyy-mm-dd")
            wshTEC_Local.Cells(rowToBeUpdated, fTECNoFacture).value = "Radiation"
        End If
    Next r
    
    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call Log_Record("modTEC_Radiation:TEC_Radiation_Update_As_Billed_Locally", "", startTime)

End Sub

Sub shp_TEC_Radiation_Impression_Click()

    Call Radiation_Apercu_Avant_Impression

End Sub

Sub Radiation_Apercu_Avant_Impression()

    Dim ws As Worksheet: Set ws = wshTEC_Radiation
    
    Dim rngToPrint As Range
    Set rngToPrint = ws.Range("C1:K35")
    
    Application.EnableEvents = False

'    'Caractères pour le rapport
'    With rngToPrint.offset(1).Font
'        .Name = "Aptos Narrow"
'        .size = 10
'    End With
'
    Application.EnableEvents = True
    
    DoEvents

    Dim header1 As String: header1 = "Radiation des TEC au  " & wshTEC_Radiation.Range("K3").value
    Dim header2 As String: header2 = wshTEC_Radiation.Range("F3").value
    
    Call Simple_Print_Setup(wshTEC_Radiation, rngToPrint, header1, header2, "$1:$1", "L")

    ws.PrintPreview
    
    'Libérer la mémoire
    Set rngToPrint = Nothing
    Set ws = Nothing
    
End Sub

Sub shp_TEC_Radiation_Back_To_TEC_Menu_Click()

    Call TEC_Radiation_Back_To_TEC_Menu
    
End Sub

Sub TEC_Radiation_Back_To_TEC_Menu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Radiation:TEC_Radiation_Back_To_TEC_Menu", "", 0)
    
    wshTEC_Radiation.Visible = xlSheetHidden
    
    fromMenu = False
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select
    
    Call Log_Record("modTEC_Radiation:TEC_Radiation_Back_To_TEC_Menu", "", startTime)

End Sub

Sub Prepare_Pour_Nouvelle_Radiation()

    Dim ws As Worksheet
    Set ws = wshTEC_Radiation
    
    With ws
        .Range("B6:B32").value = ""
        .Range("D6:K32").ClearContents
        .Range("D6:K32").Font.Bold = False
        .Shapes("Impression").Visible = False
        .Shapes("Radiation").Visible = False
        Application.EnableEvents = False
            .Range("F3").value = ""
            .Range("K3").value = ""
        Application.EnableEvents = True
        previousCellAddress = .Range("F3").Address
        .Range("F3").Select
    End With

    'Supprimer les cases à cocher existantes (au cas où)
    Dim chkBox As checkBox
    Dim headerChkBox As checkBox
    For Each chkBox In ws.CheckBoxes
        chkBox.Delete
    Next chkBox

End Sub

