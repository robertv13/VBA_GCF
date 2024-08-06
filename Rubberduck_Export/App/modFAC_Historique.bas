Attribute VB_Name = "modFAC_Historique"
Option Explicit

Sub Affiche_Liste_Factures()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("wshFAC_Historique:Affiche_Liste_Factures()")
    
    Dim ws As Worksheet: Set ws = wshFAC_Historique
    
    Application.ScreenUpdating = False
    
    Dim clientName As String: clientName = ws.Range("F4").value
    Dim dateFrom As Date: dateFrom = ws.Range("P6").value
    Dim dateTo As Date: dateTo = ws.Range("R6").value
    
    'What is the ID for the selected client ?
    Dim myInfo() As Variant
    Dim rng As Range: Set rng = wshBD_Clients.Range("dnrClients_Names_Only")
    myInfo = Fn_Find_Data_In_A_Range(rng, 1, clientName, 2)
    If myInfo(1) = "" Then
        MsgBox "Je ne peux retrouver ce client dans ma liste de clients", vbCritical
        GoTo Clean_Exit
    End If
    wshFAC_Entête.Range("W3").value = myInfo(3)
    
    Call FAC_Entête_AdvancedFilter
    ws.Range("E9:R33").ClearContents
    Call Copy_List_Of_Invoices_to_Worksheet(dateFrom, dateTo)
    
    Application.ScreenUpdating = True
    
    Call Shape_Is_Visible(False)
    
    Call Output_Timer_Results("wshFAC_Historique:Affiche_Liste_Factures()", timerStart)

Clean_Exit:
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set rng = Nothing
    Set ws = Nothing
    
End Sub

Sub FAC_Entête_AdvancedFilter() '2024-06-27 @ 15:27

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    With ws
        'Setup source data including headers
        Dim lastUsedRow As Long
        lastUsedRow = .Range("A99999").End(xlUp).row
        If lastUsedRow < 3 Then Exit Sub 'No data to filter
        Dim sourceRng As Range: Set sourceRng = .Range("A2:V" & lastUsedRow)
        
        'Define the criteria range including headers
        Dim criteriaRng As Range: Set criteriaRng = ws.Range("X2:X3")
    
        'Setup the destination Range and clear it before applying AdvancedFilter
        Dim destinationRng As Range: Set destinationRng = .Range("Z2:AU2")
        lastUsedRow = .Range("Z99999").End(xlUp).row
        If lastUsedRow > 2 Then
            ws.Range("Z3:AU" & lastUsedRow).ClearContents
        End If
    
        ' Apply the advanced filter
        sourceRng.AdvancedFilter xlFilterCopy, criteriaRng, destinationRng, False
        
        lastUsedRow = .Range("Z99999").End(xlUp).row
        If lastUsedRow < 4 Then Exit Sub
        With ws.Sort 'Sort - Inv_No
            .SortFields.clear
            .SortFields.add key:=wshTEC_Local.Range("Z3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Invoice Number
            .SetRange ws.Range("Z3:AU" & lastUsedRow) 'Set Range
            .Apply 'Apply Sort
         End With
     End With

    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set criteriaRng = Nothing
    Set destinationRng = Nothing
    Set sourceRng = Nothing
    Set ws = Nothing

End Sub

Sub Copy_List_Of_Invoices_to_Worksheet(dateMin As Date, dateMax As Date)

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    Dim ws2 As Worksheet: Set ws2 = wshENC_Détails
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("Z9999").End(xlUp).row
    If lastUsedRow < 3 Then Exit Sub 'Nothing to display
    
    Dim arr() As Variant
    ReDim arr(1 To 250, 1 To 15)
    
    With ws
        Dim i As Long, r As Long, invNo As String
        For i = 3 To lastUsedRow
            If .Range("AA" & i).value < dateMin Or .Range("AA" & i).value > dateMax Then
                GoTo nextIteration
            End If
            r = r + 1
            invNo = .Range("Z" & i).value
            arr(r, 1) = invNo
            arr(r, 2) = .Range("AA" & i).value
            arr(r, 4) = .Range("AI" & i).value
            arr(r, 6) = .Range("AK" & i).value
            arr(r, 7) = .Range("AM" & i).value
            arr(r, 8) = .Range("AO" & i).value
            arr(r, 9) = .Range("AQ" & i).value
            arr(r, 10) = .Range("AS" & i).value
            arr(r, 11) = .Range("AU" & i).value
            arr(r, 12) = .Range("AT" & i).value
            arr(r, 13) = Round(Now() - arr(r, 2), 0)
            arr(r, 14) = Fn_Get_AR_Balance_For_Invoice(ws2, invNo)
nextIteration:
        Next i
    End With
    
    If r = 0 Then
        MsgBox "Il n'y a aucune facture pour la période recherchée", vbExclamation
        GoTo Clean_Exit
    End If
    
    'Transfer the arr to the worksheet, after resizing it
    Call Array_2D_Resizer(arr, r, 14)

    With wshFAC_Historique
        For i = 1 To UBound(arr, 1)
            .Range("E" & i + 8).value = arr(i, 1)
            .Range("F" & i + 8).value = arr(i, 2)
            .Range("H" & i + 8).value = arr(i, 4)
            .Range("J" & i + 8).value = arr(i, 6)
            .Range("K" & i + 8).value = arr(i, 7)
            .Range("L" & i + 8).value = arr(i, 8)
            .Range("M" & i + 8).value = arr(i, 9)
            .Range("N" & i + 8).value = arr(i, 10)
            .Range("O" & i + 8).value = arr(i, 11)
            .Range("P" & i + 8).value = arr(i, 12)
            .Range("Q" & i + 8).value = Now() - arr(i, 2)
            .Range("R" & i + 8).value = arr(i, 12) - arr(i, 14) 'Balance
        Next i
    End With
    
    lastUsedRow = i + 8
    Call Remove_All_PDF_Icons
    If lastUsedRow >= 9 Then
        Call Insert_PDF_Icons(lastUsedRow)
    End If
Clean_Exit:

    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set ws2 = Nothing
    
End Sub

Sub Insert_PDF_Icons(lastUsedRow As Long)

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("FAC_Histo")
    
    Dim i As Long
    Dim iconPath As String
    iconPath = rootPath & Application.PathSeparator & "Resources\AdobeAcrobatReader.png"
    
    Dim pic As Picture
    Dim cell As Range
    
    'Loop through each row and insert the icon if there is data in column E
    For i = 9 To lastUsedRow
        If ws.Cells(i, 5).value <> "" Then 'Check if there is data in column E
            Set cell = ws.Cells(i, 19) 'Set the cell where the icon should be inserted (column S)
            
            'Insert the icon
            Set pic = ws.Pictures.Insert(iconPath)
            With pic
                .Top = cell.Top + 1
                .Left = cell.Left + 5
                .Height = cell.Height - 10
                .width = cell.width - 10
                .Placement = xlMoveAndSize
                .OnAction = "Display_PDF_Invoice"
            End With
        End If
    Next i
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    Set pic = Nothing
    Set ws = Nothing
    
End Sub

Sub Display_PDF_Invoice()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("FAC_Histo")
    
    Dim rowNumber As Long
    Dim fullPDFFileName As String
    
    'Determine which icon was clicked and get the corresponding row number
    Dim targetCell As Range
    Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    rowNumber = targetCell.row
    
    'Assuming the invoice number is in column E (5th column)
    fullPDFFileName = rootPath & FACT_PDF_PATH & _
        Application.PathSeparator & ws.Cells(rowNumber, 5).value & ".pdf"
    
    'Open the invoice using Adobe Acrobat Reader
    If fullPDFFileName <> "" Then
        Shell "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe " & Chr(34) & fullPDFFileName & Chr(34), vbNormalFocus
    Else
        MsgBox "Je ne retrouve pas cette facture", vbExclamation
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set targetCell = Nothing
    Set ws = Nothing
    
End Sub

Sub Remove_All_PDF_Icons() 'RMV - 2024-07-24 @ 19:58

    Dim ws As Worksheet: Set ws = wshFAC_Historique
    
    Dim pic As Picture
    For Each pic In ws.Pictures
        pic.delete
    Next pic
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set pic = Nothing
    Set ws = Nothing
    
End Sub

Sub Test_Advanced_Filter_FAC_Entête() '2024-06-27 @ 14:51

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'Clear previous results
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("Z9999").End(xlUp).row
    ws.Range("Z3:AU" & lastUsedRow).ClearContents

    'Define the source range including headers
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    Dim srcRange As Range: Set srcRange = ws.Range("A2:V" & lastUsedRow)

    'Define the criteria range including headers
    Dim criteriaRange As Range: Set criteriaRange = ws.Range("X2:X3")

    'Define the destination range starting from Y3
    Dim destRange As Range: Set destRange = ws.Range("Z2:AU2")

    'Apply the advanced filter
    srcRange.AdvancedFilter action:=xlFilterCopy, criteriaRange:=criteriaRange, CopyToRange:=destRange, Unique:=False
    
    Dim lastResultRow As Long
    lastResultRow = ws.Range("Z9999").End(xlUp).row
    If lastResultRow < 4 Then Exit Sub
    With ws.Sort 'Sort - Inv_No
        .SortFields.clear
        .SortFields.add key:=wshTEC_Local.Range("Z3"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal 'Sort Based On Invoice Number
        .SetRange ws.Range("Z3:AU" & lastResultRow) 'Set Range
        .Apply 'Apply Sort
     End With

    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set criteriaRange = Nothing
    Set destRange = Nothing
    Set srcRange = Nothing
    Set ws = Nothing
    
End Sub

Sub FAC_Historique_Clear_All_Cells()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Historique:FAC_Historique_Clear_All_Cells()")
    
    'Efface toutes les cellules de la feuille
    Application.EnableEvents = False
    ActiveSheet.Unprotect
    With wshFAC_Historique
        .Range("F4:I4,F6:I6").ClearContents
        .Range("E9:R33").ClearContents
        .Range("P6,R6").ClearContents
        Call Remove_All_PDF_Icons
        Application.EnableEvents = True
        wshFAC_Historique.Activate
        wshFAC_Historique.Range("F4").Select
    End With
    ActiveSheet.Protect UserInterfaceOnly:=True
    
    Call Output_Timer_Results("modFAC_Historique:FAC_Historique_Clear_All_Cells()", timerStart)

End Sub

Sub Shape_Is_Visible(a As Boolean)

    Dim shp As Shape: Set shp = ThisWorkbook.Sheets("FAC_Histo").Shapes("Rectangle : coins arrondis 2")
    
    If a = True Then
        shp.Visible = True
    Else
        shp.Visible = False
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set shp = Nothing
    
End Sub

Sub FAC_Historique_Back_To_FAC_Menu()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Historique:FAC_Historique_Back_To_FAC_Menu()")
    
    wshFAC_Historique.Visible = xlSheetHidden
    
    wshMenuFAC.Activate
    Call SlideIn_PrepFact
    Call SlideIn_SuiviCC
    Call SlideIn_Encaissement
    Call SlideIn_FAC_Historique
    
    wshMenuFAC.Range("A1").Select
    
    Call Output_Timer_Results("modFAC_Historique:FAC_Historique_Back_To_FAC_Menu()", timerStart)

End Sub


