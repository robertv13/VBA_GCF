Attribute VB_Name = "modFAC_Historique"
Option Explicit

Sub Affiche_Liste_Factures()

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshFAC_Historique:Affiche_Liste_Factures", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Historique
    
    Application.ScreenUpdating = False
    
    Dim clientName As String: clientName = ws.Range("D4").value
    Dim dateFrom As Date: dateFrom = ws.Range("N6").value
    Dim dateTo As Date: dateTo = ws.Range("P6").value
    
    'What is the ID for the selected client ?
    Dim myInfo() As Variant
    Dim rng As Range: Set rng = wshBD_Clients.Range("dnrClients_Names_Only")
    myInfo = Fn_Find_Data_In_A_Range(rng, 1, clientName, 2)
    If myInfo(1) = "" Then
        MsgBox "Je ne peux retrouver ce client dans ma liste de clients", vbCritical
        GoTo Clean_Exit
    End If
    wshFAC_Ent�te.Range("X3").value = myInfo(3)
    
    Call FAC_Ent�te_AdvancedFilter_Code_Client
    Application.EnableEvents = False
    ws.Range("E9:R33").ClearContents
    Application.EnableEvents = True
    Call Copy_List_Of_Invoices_to_Worksheet(dateFrom, dateTo)
    
    Application.ScreenUpdating = True
    
    Dim shp As Shape: Set shp = wshFAC_Historique.Shapes("cmdAfficheFactures")
    shp.Visible = False
    
    Call Log_Record("wshFAC_Historique:Affiche_Liste_Factures", startTime)

Clean_Exit:
    
    'Clean up
    Set rng = Nothing
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub FAC_Ent�te_AdvancedFilter_Code_Client() '2024-06-27 @ 15:27

    Dim ws As Worksheet: Set ws = wshFAC_Ent�te
    
    With ws
        'Setup the destination Range and clear it before applying AdvancedFilter
        Dim lastUsedRow As Long
        Dim destinationRng As Range: Set destinationRng = .Range("Z2:AU2")
        lastUsedRow = .Range("Z99999").End(xlUp).Row
        If lastUsedRow > 2 Then
            ws.Range("Z3:AU" & lastUsedRow).ClearContents
        End If
        
        'Setup source data including headers
        lastUsedRow = .Range("A99999").End(xlUp).Row
        If lastUsedRow < 3 Then Exit Sub 'No data to filter
        Dim sourceRng As Range: Set sourceRng = .Range("A2:V" & lastUsedRow)
        
        'Define the criteria range including headers
        Dim criteriaRng As Range: Set criteriaRng = ws.Range("X2:X3")
    
        ' Apply the advanced filter
        sourceRng.AdvancedFilter xlFilterCopy, criteriaRng, destinationRng, False
        
        lastUsedRow = .Range("Z99999").End(xlUp).Row
        If lastUsedRow < 4 Then Exit Sub
        With ws.Sort 'Sort - Inv_No
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("Z3"), _
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

Sub FAC_Ent�te_AdvancedFilter_AC_C() '2024-07-19 @ 13:58

    Dim ws As Worksheet: Set ws = wshFAC_Ent�te
    
    With ws
        'Setup the destination Range and clear it before applying AdvancedFilter
        Dim lastUsedRow As Long
        Dim destinationRng As Range: Set destinationRng = .Range("AY2:BP2")
        lastUsedRow = ws.Cells(ws.rows.count, "AY").End(xlUp).Row
        If lastUsedRow > 2 Then
            ws.Range("AY3:BP" & lastUsedRow).ClearContents
        End If
        
        'Setup source data including headers
        lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
        If lastUsedRow < 3 Then Exit Sub 'No data to filter
        Dim sourceRng As Range: Set sourceRng = .Range("A2:V" & lastUsedRow)
        
        'Define the criteria range including headers
        Dim criteriaRng As Range: Set criteriaRng = ws.Range("AW2:AW3")
    
        ' Apply the advanced filter
        sourceRng.AdvancedFilter xlFilterCopy, criteriaRng, destinationRng, False
        
        lastUsedRow = ws.Cells(ws.rows.count, "AY").End(xlUp).Row
        If lastUsedRow < 4 Then Exit Sub
        With ws.Sort 'Sort - Inv_No
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("AY3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Invoice Number
            .SetRange ws.Range("AY3:BP" & lastUsedRow) 'Set Range
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

    Dim ws As Worksheet: Set ws = wshFAC_Ent�te
    Dim ws2 As Worksheet: Set ws2 = wshENC_D�tails
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("Z9999").End(xlUp).Row
    If lastUsedRow < 3 Then Exit Sub 'Nothing to display
    
    Dim arr() As Variant
    ReDim arr(1 To 250, 1 To 11)
    
    With ws
        Dim i As Long, r As Long
        For i = 3 To lastUsedRow
            If .Range("AA" & i).value >= dateMin And .Range("AA" & i).value <= dateMax And _
                .Range("AB" & i).value = "C" Then
                r = r + 1
                arr(r, 1) = .Range("Z" & i).value  'Invoice number
                arr(r, 2) = .Range("AA" & i).value 'Invoice Date
                arr(r, 3) = .Range("AI" & i).value 'Fees
                arr(r, 4) = .Range("AK" & i).value 'Misc. 1
                arr(r, 5) = .Range("AM" & i).value 'Misc. 2
                arr(r, 6) = .Range("AO" & i).value 'Misc. 3
                arr(r, 7) = .Range("AQ" & i).value 'GST $
                arr(r, 8) = .Range("AS" & i).value 'PST $
                arr(r, 9) = .Range("AU" & i).value 'Deposit
                arr(r, 10) = .Range("AT" & i).value 'AR_Total
                arr(r, 11) = Fn_Get_AR_Balance_For_Invoice(ws2, .Range("Z" & i).value)
            End If
        Next i
    End With
    
    If r = 0 Then
        MsgBox "Il n'y a aucune facture pour la p�riode recherch�e", vbExclamation
        GoTo Clean_Exit
    End If
    
    'Transfer the arr to the worksheet, after resizing it
    Call Array_2D_Resizer(arr, r, 14)

    Application.EnableEvents = False
    
    With wshFAC_Historique
        For i = 1 To UBound(arr, 1)
            .Range("C" & i + 8).value = arr(i, 1)
            .Range("D" & i + 8).value = arr(i, 2)
            .Range("F" & i + 8).value = arr(i, 3)
            .Range("H" & i + 8).value = arr(i, 4)
            .Range("I" & i + 8).value = arr(i, 5)
            .Range("J" & i + 8).value = arr(i, 6)
            .Range("K" & i + 8).value = arr(i, 7)
            .Range("L" & i + 8).value = arr(i, 8)
            .Range("M" & i + 8).value = arr(i, 9)
            .Range("N" & i + 8).value = arr(i, 10)
            .Range("O" & i + 8).value = Now() - arr(i, 2)
            .Range("P" & i + 8).value = arr(i, 10) - arr(i, 9) 'Balance
        Next i
    End With
    
    lastUsedRow = i + 8
    Call Remove_All_PDF_Icons
    If lastUsedRow >= 9 Then
        Call Insert_PDF_Icons(lastUsedRow)
    End If
    
    Application.EnableEvents = True

Clean_Exit:

    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set ws2 = Nothing
    
End Sub

Sub Insert_PDF_Icons(lastUsedRow As Long)

    Dim ws As Worksheet: Set ws = wshFAC_Historique
    
    Dim i As Long
    Dim iconPath As String
    iconPath = wshAdmin.Range("F5").value & Application.PathSeparator & "Resources\AdobeAcrobatReader.png"
    
    Dim pic As Picture
    Dim cell As Range
    
    'Loop through each row and insert the icon if there is data in column E
    For i = 9 To lastUsedRow
        If ws.Cells(i, 3).value <> "" Then 'Check if there is data in column C
            Set cell = ws.Cells(i, 17) 'Set the cell where the icon should be inserted (column Q)
            
            'Insert the icon
            Set pic = ws.Pictures.Insert(iconPath)
            Debug.Print pic.width, pic.Height
            With pic
                .Top = cell.Top + 1
                .Left = cell.Left + 5
                .Height = cell.Height - 15
                .width = cell.width - 15
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

    Dim ws As Worksheet: Set ws = wshFAC_Historique
    
    Dim rowNumber As Long
    Dim fullPDFFileName As String
    
    'Determine which icon was clicked and get the corresponding row number
    Dim targetCell As Range
    Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    rowNumber = targetCell.Row
    
    'Assuming the invoice number is in column E (5th column)
    fullPDFFileName = wshAdmin.Range("F5").value & FACT_PDF_PATH & _
        Application.PathSeparator & ws.Cells(rowNumber, 3).value & ".pdf"
    
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
        pic.Delete
    Next pic
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set pic = Nothing
    Set ws = Nothing
    
End Sub

Sub Test_Advanced_Filter_FAC_Ent�te() '2024-06-27 @ 14:51

    Dim ws As Worksheet: Set ws = wshFAC_Ent�te
    
    'Clear previous results
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("Z9999").End(xlUp).Row
    ws.Range("Z3:AU" & lastUsedRow).ClearContents

    'Define the source range including headers
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    Dim srcRange As Range: Set srcRange = ws.Range("A2:V" & lastUsedRow)

    'Define the criteria range including headers
    Dim criteriaRange As Range: Set criteriaRange = ws.Range("X2:X3")

    'Define the destination range starting from Y3
    Dim destRange As Range: Set destRange = ws.Range("Z2:AU2")

    'Apply the advanced filter
    srcRange.AdvancedFilter action:=xlFilterCopy, _
                            criteriaRange:=criteriaRange, _
                            CopyToRange:=destRange, _
                            Unique:=False
    
    Dim lastResultRow As Long
    lastResultRow = ws.Range("Z9999").End(xlUp).Row
    If lastResultRow < 4 Then Exit Sub
    With ws.Sort 'Sort - Inv_No
        .SortFields.Clear
        .SortFields.Add key:=wshTEC_Local.Range("Z3"), _
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

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Historique:FAC_Historique_Clear_All_Cells", 0)
    
    'Efface toutes les cellules de la feuille
    Application.EnableEvents = False
    ActiveSheet.Unprotect
    With wshFAC_Historique
        .Range("D4:H4", "D6:F6").ClearContents
        .Range("E9:R33").ClearContents
        .Range("P6,R6").ClearContents
        Call Remove_All_PDF_Icons
        Application.EnableEvents = True
        wshFAC_Historique.Activate
        wshFAC_Historique.Range("F4").Select
    End With
    
    With ActiveSheet
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

    Call Log_Record("modFAC_Historique:FAC_Historique_Clear_All_Cells", startTime)

End Sub

Sub FAC_Historique_Back_To_FAC_Menu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Historique:FAC_Historique_Back_To_FAC_Menu", 0)
    
    wshFAC_Historique.Visible = xlSheetHidden
    
'    Call SlideIn_PrepFact
'    Call SlideIn_SuiviCC
'    Call SlideIn_Encaissement
'    Call SlideIn_FAC_Historique
    
    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select
    
    Call Log_Record("modFAC_Historique:FAC_Historique_Back_To_FAC_Menu", startTime)

End Sub


