Attribute VB_Name = "modTEC_Radiation"
Option Explicit

Sub TEC_Radiation_Procedure(codeClient As String, cutoffDate As String)

    If cutoffDate = vbNullString Then
        Exit Sub
    End If
        
    Call modImport.ImporterTEC
    
    Dim maxDate As Date
    Dim Y As Integer, m As Integer, d As Integer
    Y = year(cutoffDate)
    m = month(cutoffDate)
    d = day(cutoffDate)
    maxDate = DateSerial(Y, m, d)
    
    Dim ws As Worksheet: Set ws = wshTEC_Radiation
    Dim wsSource As Worksheet: Set wsSource = wsdTEC_Local
    
    'AdvancedFilter # 2 avec TEC_Local (Heures Facturables, Non Facturées, Non Détruites à la date Limite)
    Call Get_TEC_For_Client_AF(codeClient, CDate(cutoffDate), "VRAI", "FAUX", "FAUX")
    
    'Avons-nous des résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, "AQ").End(xlUp).Row
    If lastUsedRow < 3 Then
        MsgBox "Il n'y a aucun TEC pour ce client", vbInformation
        Call Prepare_Pour_Nouvelle_Radiation
        wshTEC_Radiation.Range("F3").Activate
        GoTo exitSub
    End If
    
    'Transfère la table en mémoire (arr)
    Dim arr As Variant
    arr = wsSource.Range("AQ3:BF" & lastUsedRow).Value
    
    Dim i As Long
    Dim tecID As Long
    Dim dateTEC As Date
    Dim profInit As String, descTEC As String
    Dim hresTEC As Currency, tauxHoraire As Currency, valeurTEC As Currency
    Dim totalHresTEC As Currency, totalValeurTEC As Currency
    Dim currRow As Integer, activeRow As Long
    Dim vueIncomplete  As Boolean
    vueIncomplete = False
    currRow = 6
    For i = 1 To UBound(arr, 1)
        Debug.Print currRow
        If currRow <= 30 Then
            tecID = CLng(arr(i, fTECTECID))
            dateTEC = Format$(arr(i, fTECDate), wsdADMIN.Range("USER_DATE_FORMAT").Value)
            profInit = Format$(arr(i, fTECProfID), "000") & arr(i, fTECProf)
            descTEC = arr(i, fTECDescription)
            hresTEC = arr(i, fTECHeures)
            With ws
                .Cells(currRow, 5).Value = tecID
                .Cells(currRow, 6).Value = dateTEC
                .Cells(currRow, 7).Value = Mid$(profInit, 4)
                .Cells(currRow, 8).Value = descTEC
                .Cells(currRow, 10).Value = hresTEC
                tauxHoraire = Fn_Get_Hourly_Rate(CLng(Left$(profInit, 3)), CDate(cutoffDate))
                valeurTEC = hresTEC * tauxHoraire
                .Cells(currRow, 11).Value = valeurTEC
            End With
            activeRow = currRow
            totalHresTEC = totalHresTEC + hresTEC
            totalValeurTEC = totalValeurTEC + valeurTEC
        Else
            vueIncomplete = True
        End If
        currRow = currRow + 1
    Next i
    
    'La ligne maximum ne doit pas excéder 32
    currRow = currRow + 1
    If currRow > 32 Then
        currRow = 32
    End If
    
    If vueIncomplete Then
        MsgBox _
            Prompt:="L'affichage des heures n'est pas complet", _
            Title:="Maximum de 25 lignes sont affichées", _
            Buttons:=vbInformation
    End If
    
    'Affiche les totaux
    With ws
        .Cells(3, 9).Value = "Total heures TEC = " & Format$(totalHresTEC, "#,##0.00")
        .Cells(currRow, 8).Value = "* TOTAUX des TEC qui seront RADIÉS *"
        .Cells(currRow, 8).Font.Bold = True
        .Cells(currRow, 10).Font.Bold = True
        .Cells(currRow, 11).Font.Bold = True
    End With
    
    Call AjouterCheckBoxesAvecControleGlobal(activeRow)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ws.Shapes("Impression").Visible = True
    ws.Shapes("Radiation").Visible = False
    
exitSub:

    'Libérer la mémoire
    Set ws = Nothing
    Set wsSource = Nothing
    
End Sub

Sub AjouterCheckBoxesAvecControleGlobal(lastUsedRow As Long)

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
        .Caption = vbNullString
        .OnAction = "CocherOuDecocherToutesCases" 'Associe la macro de contrôle global
    End With

    'Ajouter une case à cocher pour chaque ligne du tableau
    For i = 6 To lastUsedRow
        Set chkBox = ws.CheckBoxes.Add(Left:=ws.Cells(i, 4).Left + 5, _
                                       Top:=ws.Cells(i, 4).Top, _
                                       Width:=ws.Cells(i, 4).Width, _
                                       Height:=ws.Cells(i, 4).Height)
        With chkBox
            .Name = "chk_" & i
            .Caption = vbNullString
            .linkedCell = ws.Cells(i, 2).Address
            .OnAction = "CalculerTotaux"
        End With
    Next i

End Sub

Sub CocherOuDecocherToutesCases()

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
            If chkBox.Value <> xlOn Then
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
    headerChkBox.Value = xlOff 'On s'assure que la case est décochée avant de basculer
    headerChkBox.Value = IIf(newState, xlOn, xlOff) 'Appliquer l'état approprié
    
    'Parcourir toutes les cases à cocher et appliquer le nouvel état
    For Each chkBox In ws.CheckBoxes
        If chkBox.Name <> "chk_header" Then
            chkBox.Value = IIf(newState, xlOn, xlOff)
            chkBox.Characters.text = vbNullString
            chkBox.Characters.Caption = vbNullString
        End If
    Next chkBox

    'Réactiver les événements
    Application.EnableEvents = True

    'Protéger la feuille à nouveau
    ws.Protect Password:=vbNullString, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
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
            If chkBox.Value = xlOn Then
                'Ajouter la valeur de la cellule correspondante à la somme totale
                totalHres = totalHres + ws.Cells(rowNum, "J").Value
                totalValeur = totalValeur + ws.Cells(rowNum, "K").Value
                
            End If
            rowNum = rowNum + 1
        End If
    Next chkBox
    
    ' Afficher le total dans une cellule spécifique (par exemple, cellule Z1)
    ws.Cells(rowNum + 1, 10).Value = Format$(totalHres, "###,##0.00")
    ws.Cells(rowNum + 1, 11).Value = Format$(totalValeur, "###,##0.00 $")
    
    If Not totalHres = 0 Then
        ws.Shapes("Radiation").Visible = True
    Else
        ws.Shapes("Radiation").Visible = False
    End If
    
End Sub

Sub shp_TEC_Radiation_GO_Click()

    Call Radiation_Mise_À_Jour
    
    Call Prepare_Pour_Nouvelle_Radiation
    
    wshTEC_Radiation.Activate

End Sub

Sub Radiation_Mise_À_Jour()

    'Avons-nous des résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "AQ").End(xlUp).Row
    
    If lastUsedRow >= 3 Then
        Call MettreAJourEstFactureeBDMaster(3, lastUsedRow)
        Call MettreAJourEstFactureeBDLocale(3, lastUsedRow)
        Call modTEC_TDB.ActualiserTECTableauDeBord
        MsgBox "Les TEC sélectionnés ont été radié avec succès", vbOKOnly, "Confirmation de traitement"
    Else
        MsgBox "AUCUNE radiation n'a été effectuée", vbOKOnly
    End If
    
End Sub

Sub MettreAJourEstFactureeBDMaster(firstRow As Long, lastRow As Long) 'Update Billed Status in DB

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Radiation:MettreAJourEstFactureeBDMaster", firstRow & ", " & lastRow, 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "TEC_Local$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    Dim r As Long, tecID As Long, strSQL As String
    'Offset entre TEC_Local & wshTEC_Radiation
    Dim offset As Long
    offset = 3
    For r = firstRow To lastRow
        If wshTEC_Radiation.Range("B" & r + offset).Value = True Then
            tecID = wsdTEC_Local.Range("AQ" & r).Value
            
            'Open the recordset for the specified ID
            strSQL = "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & tecID
            recSet.Open strSQL, conn, 2, 3
            If Not recSet.EOF Then
                'Update EstFacturee, DateFacturee & NoFacture
                recSet.Fields(fTECEstFacturee - 1).Value = "VRAI"
                recSet.Fields(fTECDateFacturee - 1).Value = Format$(Date, "yyyy-mm-dd")
                recSet.Fields(fTECNoFacture - 1).Value = "Radiation"
                recSet.Update
            Else
                'Handle the case where the specified ID is not found
                MsgBox "L'enregistrement avec le TECID '" & r & "' ne peut être trouvé!", _
                    vbExclamation
                recSet.Close
                conn.Close
                Exit Sub
            End If
            'Update the recordset (create the record)
            recSet.Update
            recSet.Close
        End If
next_iteration:
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Radiation:MettreAJourEstFactureeBDMaster", vbNullString, startTime)

End Sub

Sub MettreAJourEstFactureeBDLocale(firstResultRow As Long, lastResultRow As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Radiation:MettreAJourEstFactureeBDLocale", firstResultRow & ", " & lastResultRow, 0)
    
    'Set the range to look for
    Dim lookupRange As Range: Set lookupRange = wsdTEC_Local.Range("l_tbl_TEC_Local[TECID]")
    
    Dim r As Long, rowToBeUpdated As Long, tecID As Long
    'Offset entre TEC_Local & wshTEC_Radiation
    Dim offset As Long
    offset = 3
    For r = firstResultRow To lastResultRow
        If wshTEC_Radiation.Range("B" & r + offset).Value = True Then
            tecID = wsdTEC_Local.Range("AQ" & r).Value
            rowToBeUpdated = Fn_Find_Row_Number_TECID(tecID, lookupRange)
            wsdTEC_Local.Cells(rowToBeUpdated, fTECEstFacturee).Value = "VRAI"
            wsdTEC_Local.Cells(rowToBeUpdated, fTECDateFacturee).Value = Format$(Date, "yyyy-mm-dd")
            wsdTEC_Local.Cells(rowToBeUpdated, fTECNoFacture).Value = "Radiation"
        End If
    Next r
    
    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Radiation:MettreAJourEstFactureeBDLocale", vbNullString, startTime)

End Sub

Sub shp_TEC_Radiation_Impression_Click()

    Call Radiation_Apercu_Avant_Impression

End Sub

Sub Radiation_Apercu_Avant_Impression()

    Dim ws As Worksheet: Set ws = wshTEC_Radiation
    
    Dim shp As Shape
    For Each shp In ws.Shapes
        Debug.Print shp.Name, shp.Type, shp.Width, shp.Left
        If shp.Name = "Drop Down 193" Then shp.Delete
    Next shp
    
    Dim rngToPrint As Range
    Set rngToPrint = ws.Range("C1:L35")
    
    Dim header1 As String: header1 = "Radiation des TEC au  " & wshTEC_Radiation.Range("K3").Value
    Dim header2 As String: header2 = wshTEC_Radiation.Range("F3").Value
    
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wshTEC_Radiation, rngToPrint, header1, header2, "$1:$1", "L")

    ws.PrintPreview
    
    'Libérer la mémoire
    Set rngToPrint = Nothing
    Set ws = Nothing
    
End Sub

Sub shpRetourMenuTEC_Click()

    Call RetourMenuTEC
    
End Sub

Sub RetourMenuTEC()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Radiation:RetourMenuTEC", vbNullString, 0)
    
    wshTEC_Radiation.Visible = xlSheetHidden
    
    gFromMenu = False
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Radiation:RetourMenuTEC", vbNullString, startTime)

End Sub

Sub Prepare_Pour_Nouvelle_Radiation()

    Dim ws As Worksheet
    Set ws = wshTEC_Radiation
    
    With ws
        .Range("B6:B32").Value = vbNullString
        .Range("D6:K32").ClearContents
        .Range("D6:K32").Font.Bold = False
        .Shapes("Impression").Visible = False
        .Shapes("Radiation").Visible = False
        Application.EnableEvents = False
            .Range("F3").Value = vbNullString
            .Range("K3").Value = vbNullString
        Application.EnableEvents = True
        gPreviousCellAddress = .Range("F3").Address
        .Range("F3").Select
    End With

    'Supprimer les cases à cocher existantes (au cas où)
    Dim chkBox As checkBox
    Dim headerChkBox As checkBox
    For Each chkBox In ws.CheckBoxes
        chkBox.Delete
    Next chkBox

End Sub


