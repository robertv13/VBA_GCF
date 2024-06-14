Attribute VB_Name = "modFAC_Brouillon"
Option Explicit
Dim invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Dim lastRow As Long, lastResultRow As Long, resultRow As Long

Sub FAC_Brouillon_New_Invoice() 'Clear contents
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_New_Invoice()")
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If wshFAC_Brouillon.Range("B27").value = False Then
        With wshFAC_Brouillon
            .Range("B24").value = True
            .Range("K3:L7,O3,O5").Clearcontents 'Clear cells for a new Invoice
            .Range("J8:Q46").Clearcontents
            .Range("O6").value = .Range("FACNextInvoiceNumber").value 'Paste Invoice ID
            .Range("FACNextInvoiceNumber").value = .Range("FACNextInvoiceNumber").value + 1 'Increment Next Invoice ID
            
            Call FAC_Brouillon_Setup_All_Cells
            
            .Range("B20").value = ""
            .Range("B24").value = False
            .Range("B26").value = False
            .Range("B27").value = True 'Set the value to TRUE
        End With
        
        With wshFAC_Finale
            .Range("B21,B23:C27,E28").Clearcontents
            .Range("A34:F68").Clearcontents
            .Range("E28").value = wshFAC_Brouillon.Range("O6").value 'Invoice #
            .Range("B69:F81").Clearcontents 'NOT the formulas
            
            Call FAC_Finale_Setup_All_Cells
        
        End With
        
        wshFAC_Brouillon.Range("B16").value = False '2024-03-14 @ 08:41
        
        Call FAC_Brouillon_Clear_All_TEC_Displayed
        
        'Move on to CLient Name
        wshFAC_Brouillon.Range("E4:F4").Clearcontents
        With wshFAC_Brouillon.Range("E4:F4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        wshFAC_Brouillon.Select
        wshFAC_Brouillon.Range("E4").Select 'Start inputing values for a NEW invoice
    End If

    'Save button is disabled UNTIL the invoice is saved
    Call FAC_Finale_Disable_Save_Button
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_New_Invoice()", timerStart)

End Sub

Sub FAC_Brouillon_Client_Change(ClientName As String)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_Client_Change()")
    
    Dim myInfo() As Variant
    Dim rng As Range: Set rng = wshBD_Clients.Range("dnrClients_Names_Only")
    
    myInfo = Fn_Find_Data_In_A_Range(rng, 1, ClientName, 3)
    
    If myInfo(1) = "" Then
        MsgBox "Je ne peux retrouver ce client dans ma liste", vbCritical
        GoTo Clean_Exit
    End If
        
    wshFAC_Brouillon.Range("B18").value = wshBD_Clients.Cells(myInfo(2), 2)
    
    With wshFAC_Brouillon
        ActiveSheet.Unprotect
        .Range("K3").value = wshBD_Clients.Cells(myInfo(2), 3)
        .Range("K4").value = ClientName
        .Range("K5").value = wshBD_Clients.Cells(myInfo(2), 6) 'Adresse1
        If wshBD_Clients.Cells(myInfo(2), 7) <> "" Then
            .Range("K6").value = wshBD_Clients.Cells(myInfo(2), 7) 'Adresse2
            .Range("K7").value = wshBD_Clients.Cells(myInfo(2), 8) & " " & _
                                wshBD_Clients.Cells(myInfo(2), 9) & "  " & _
                                wshBD_Clients.Cells(myInfo(2), 10) 'Ville, Province & Code postal
        Else
            .Range("K6").value = wshBD_Clients.Cells(myInfo(2), 8) & " " & _
                                wshBD_Clients.Cells(myInfo(2), 9) & "  " & _
                                wshBD_Clients.Cells(myInfo(2), 10) 'Ville, Province & Code postal
            .Range("K7").value = ""
        End If
    End With
    
    With wshFAC_Finale
        .Range("B23").value = wshBD_Clients.Cells(myInfo(2), 3)
        .Range("B24").value = ClientName
        .Range("B25").value = wshBD_Clients.Cells(myInfo(2), 6) 'Adresse1
        If wshBD_Clients.Cells(myInfo(2), 7) <> "" Then
            .Range("B26").value = wshBD_Clients.Cells(myInfo(2), 7) 'Adresse2
            .Range("B27").value = wshBD_Clients.Cells(myInfo(2), 8) & " " & _
                                wshBD_Clients.Cells(myInfo(2), 9) & "  " & _
                                wshBD_Clients.Cells(myInfo(2), 10) 'Ville, Province & Code postal
        Else
            .Range("B26").value = wshBD_Clients.Cells(myInfo(2), 8) & " " & _
                                wshBD_Clients.Cells(myInfo(2), 9) & "  " & _
                                wshBD_Clients.Cells(myInfo(2), 10) 'Ville, Province & Code postal
            .Range("B27").value = ""
        End If
    End With
    
    Call FAC_Brouillon_Clear_All_TEC_Displayed
    
    wshFAC_Brouillon.Range("O3").Select 'Move on to Invoice Date

Clean_Exit:

    Set rng = Nothing
    
    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_Client_Change()", timerStart)
    
End Sub

Sub FAC_Brouillon_Date_Change(d As String)

    Application.EnableEvents = False
    
    If InStr(wshFAC_Brouillon.Range("O6").value, "-") = 0 Then
        Dim y As String
        y = Right(Year(d), 2)
        wshFAC_Brouillon.Range("O6").value = y & "-" & wshFAC_Brouillon.Range("O6").value
        wshFAC_Finale.Range("E28").value = wshFAC_Brouillon.Range("O6").value
    End If
    
    wshFAC_Finale.Range("B21").value = "Le " & Format(d, "d mmmm yyyy")
    
    'Must Get GST & PST rates and store them in wshFAC_Brouillon 'B' column
    Dim DateTaxRates As Date
    DateTaxRates = d
    wshFAC_Brouillon.Range("B29").value = Fn_Get_Tax_Rate(DateTaxRates, "F")
    wshFAC_Brouillon.Range("B30").value = Fn_Get_Tax_Rate(DateTaxRates, "P")
        
    Dim cutoffDate As Date
    cutoffDate = d
    Call FAC_Brouillon_Get_All_TEC_By_Client(cutoffDate, False)
    
    Dim rng As Range
    Set rng = wshFAC_Brouillon.Range("O3")
    Call Fill_Or_Empty_Range_Background(rng, False)
    
    Set rng = wshFAC_Brouillon.Range("L11")
'    Call Fill_Or_Empty_Range_Background(rng, True, 6)

    On Error Resume Next
    wshFAC_Brouillon.Range("L11").Select 'Move on to Services Entry
    On Error GoTo 0
    
    Application.EnableEvents = True
    
End Sub

Sub FAC_Brouillon_Inclure_TEC_Factures_Click()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_Inclure_TEC_Factures_Click()")
    
    Dim cutoffDate As Date
    cutoffDate = wshFAC_Brouillon.Range("O3").value
    
    If wshFAC_Brouillon.Range("B16").value = True Then
        Call FAC_Brouillon_Get_All_TEC_By_Client(cutoffDate, True)
    Else
        Call FAC_Brouillon_Get_All_TEC_By_Client(cutoffDate, False)
    End If
    
    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_Inclure_TEC_Factures_Click()", timerStart)

End Sub

Sub FAC_Brouillon_Setup_All_Cells()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_Setup_All_Cells()")

    Application.EnableEvents = False
    
    Dim rng As Range
    
    With wshFAC_Brouillon
        Set rng = .Range("L11:O45")
        rng.Clearcontents
        Call Fill_Or_Empty_Range_Background(rng, False)

        .Range("J47:P60").Clearcontents
        
        Call FAC_Brouillon_Set_Labels(.Range("K47"), "FAC_Label_SubTotal_1")
        Call FAC_Brouillon_Set_Labels(.Range("K51"), "FAC_Label_SubTotal_2")
        Call FAC_Brouillon_Set_Labels(.Range("K52"), "FAC_Label_TPS")
        Call FAC_Brouillon_Set_Labels(.Range("K53"), "FAC_Label_TVQ")
        Call FAC_Brouillon_Set_Labels(.Range("K55"), "FAC_Label_GrandTotal")
        Call FAC_Brouillon_Set_Labels(.Range("K57"), "FAC_Label_Deposit")
        Call FAC_Brouillon_Set_Labels(.Range("K59"), "FAC_Label_AmountDue")
        
        .Range("M47").formula = "=IF(SUM(M11:M45),SUM(M11:M45),B19)"   'Total hours entered OR TEC selected"
        .Range("N47").formula = wshAdmin.Range("TauxHoraireFacturation") 'Rate per hour
        .Range("O47").formula = "=M47*N47"                               'Fees sub-total
        .Range("O47").Font.Bold = True
        
        .Range("M48").value = wshAdmin.Range("FAC_Label_Frais_1").value 'Misc. # 1 - Descr.
        .Range("O48").value = ""                                        'Misc. # 1 - Amount
        .Range("M49").value = wshAdmin.Range("FAC_Label_Frais_2").value 'Misc. # 2 - Descr.
        .Range("O49").value = ""                                        'Misc. # 2 - Amount
        .Range("M50").value = wshAdmin.Range("FAC_Label_Frais_3").value 'Misc. # 3 - Descr.
        .Range("O50").value = ""                                        'Misc. # 3 - Amount
        
        .Range("O51").formula = "=sum(O47:O50)"                         'Sub-total
        .Range("O51").Font.Bold = True
        
        .Range("N52").value = wshFAC_Brouillon.Range("B29").value       'GST Rate
        .Range("N52").NumberFormat = "0.00%"
        .Range("O52").formula = "=round(o51*n52,2)"                     'GST Amnt
        .Range("N53").value = wshFAC_Brouillon.Range("B30").value       'PST Rate
        .Range("N53").NumberFormat = "0.000%"
        .Range("O53").formula = "=round(o51*n53,2)"                     'PST Amnt
        .Range("O55").formula = "=sum(o51:o54)"                         'Grand Total"
        .Range("O57").value = ""
        .Range("O59").formula = "=O55-O57"                              'Deposit Amount
        
    End With
    
    Application.EnableEvents = True
    
    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_Setup_All_Cells()", timerStart)

End Sub

Sub FAC_Brouillon_Set_Labels(r As Range, l As String)

    r.value = wshAdmin.Range(l).value
    If wshAdmin.Range(l & "_Bold").value = "OUI" Then r.Font.Bold = True

End Sub

Sub FAC_Brouillon_Goto_Misc_Charges()
    
    ActiveWindow.SmallScroll Down:=6
    wshFAC_Brouillon.Range("M47").Select 'Hours Summary
    
End Sub

Sub FAC_Brouillon_Clear_All_TEC_Displayed()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_Clear_All_TEC_Displayed()")
    
    Application.EnableEvents = False
    
    Dim lastRow As Long
    lastRow = wshFAC_Brouillon.Range("D9999").End(xlUp).row
    If lastRow > 7 Then
        wshFAC_Brouillon.Range("D8:I" & lastRow + 2).Clearcontents
        Call FAC_Brouillon_TEC_Remove_Check_Boxes(lastRow)
    End If
    
    Application.EnableEvents = True

    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_Clear_All_TEC_Displayed()", timerStart)

End Sub

Sub FAC_Brouillon_Get_All_TEC_By_Client(d As Date, includeBilledTEC As Boolean)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_Get_All_TEC_By_Client()")
    
    'Set all criteria before calling FAC_Brouillon_TEC_Advanced_Filter_And_Sort
    Dim c1 As Long, c2 As String, c3 As Boolean
    Dim c4 As Boolean, c5 As Boolean
    c1 = wshFAC_Brouillon.Range("B18").value
    c2 = "<=" & Format(d, "mm-dd-yyyy")
    c3 = True
    If includeBilledTEC Then c4 = True Else c4 = False
    c5 = False

    Call FAC_Brouillon_Clear_All_TEC_Displayed
    Call FAC_Brouillon_TEC_Advanced_Filter_And_Sort(c1, c2, c3, c4, c5)
    Call FAC_Brouillon_TEC_Filtered_Entries_Copy_To_FAC_Brouillon
    
    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_Get_All_TEC_By_Client()", timerStart)

End Sub

Sub FAC_Brouillon_TEC_Advanced_Filter_And_Sort(clientID As Long, _
        cutoffDate As String, _
        isBillable As Boolean, _
        isInvoiced As Boolean, _
        isDeleted As Boolean)
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_TEC_Advanced_Filter_And_Sort()")
    
    Application.ScreenUpdating = False

    With wshTEC_Local
        'Is there anything to filter ?
        Dim lastSourceRow As Long, lastResultRow As Long
        lastSourceRow = .Range("A99999").End(xlUp).row 'Last TEC Entry row
        If lastSourceRow < 3 Then Exit Sub 'Nothing to filter
        
        'Clear the filtered rows area
        lastResultRow = .Range("AT9999").End(xlUp).row
        If lastResultRow > 2 Then .Range("AT3:BH" & lastResultRow).Clearcontents
        
        Dim rngSource As Range, rngCriteria As Range, rngCopyToRange As Range
        Set rngSource = wshTEC_Local.Range("A2:P" & lastSourceRow)
        If clientID <> 0 Then .Range("AN3").value = clientID
        .Range("AO3").value = cutoffDate
        .Range("AP3").value = isBillable
        If isInvoiced <> True Then
            .Range("AQ3").value = isInvoiced
        Else
            .Range("AQ3").value = ""
        End If
        .Range("AR3").value = isDeleted
        Set rngCriteria = .Range("AN2:AR3")
        Set rngCopyToRange = .Range("AT2:BH2")
        
        rngSource.AdvancedFilter xlFilterCopy, rngCriteria, rngCopyToRange, Unique:=True
        
        lastResultRow = .Range("AT9999").End(xlUp).row
        If lastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort
            .SortFields.clear
            .SortFields.add key:=wshTEC_Local.Range("AW3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.add key:=wshTEC_Local.Range("AU3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Prof_ID
            .SortFields.add key:=wshTEC_Local.Range("AT3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SetRange wshTEC_Local.Range("AT3:BH" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
No_Sort_Required:
    End With
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_TEC_Advanced_Filter_And_Sort()", timerStart)

End Sub

Sub FAC_Brouillon_TEC_Filtered_Entries_Copy_To_FAC_Brouillon() '2024-03-21 @ 07:10

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_TEC_Filtered_Entries_Copy_To_FAC_Brouillon()")

    Dim lastUsedRow As Long
    lastUsedRow = wshTEC_Local.Range("AT9999").End(xlUp).row
    If lastUsedRow < 3 Then Exit Sub 'No rows
    
    Application.ScreenUpdating = False
    
    Dim arr() As Variant, totalHres As Double
    ReDim arr(1 To (lastUsedRow - 2), 1 To 6) As Variant
    With wshTEC_Local
        Dim i As Integer
        For i = 3 To lastUsedRow
            arr(i - 2, 1) = .Range("AW" & i).value 'Date
            arr(i - 2, 2) = .Range("AV" & i).value 'Prof
            arr(i - 2, 3) = .Range("AY" & i).value 'Description
            arr(i - 2, 4) = .Range("AZ" & i).value 'Heures
            totalHres = totalHres + .Range("AZ" & i).value
            arr(i - 2, 5) = .Range("BD" & i).value 'Facturée ou pas
            arr(i - 2, 6) = .Range("AT" & i).value 'TEC_ID
        Next i
        'Copy array to worksheet
        Dim rng As Range
        'Set rng = .Range("D8").Resize(UBound(arr, 1), UBound(arr, 2))
        Set rng = wshFAC_Brouillon.Range("D8").Resize(lastUsedRow - 2, UBound(arr, 2))
        rng.value = arr
    End With
    
    With wshFAC_Brouillon
        .Range("D8:H" & lastRow + 7).Font.Color = vbBlack
        .Range("D8:H" & lastRow + 7).Font.Bold = False
        
        .Range("G" & lastUsedRow + 7).value = totalHres
        .Range("G8:G" & lastUsedRow + 7).NumberFormat = "##0.00"
    End With
        
    Call FAC_Brouillon_TEC_Add_Check_Boxes(lastUsedRow)

    Application.ScreenUpdating = True

    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_TEC_Filtered_Entries_Copy_To_FAC_Brouillon()", timerStart)
    
End Sub
 
Sub FAC_Brouillon_Goto_Onglet_FAC_Finale()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_Goto_Onglet_FAC_Finale()")
   
    Application.ScreenUpdating = False
    
    Call FAC_Finale_Cacher_Heures
    Call FAC_Finale_Cacher_Sommaire_Heures
    
    wshFAC_Finale.Visible = xlSheetVisible
    wshFAC_Finale.Activate
    wshFAC_Finale.Range("I50").Select
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_Goto_Onglet_FAC_Finale()", timerStart)

End Sub

Sub FAC_Brouillon_Back_To_FAC_Menu()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_Back_To_FAC_Menu()")
   
    wshMenuFACT.Activate
    Call SlideIn_PrepFact
    Call SlideIn_SuiviCC
    Call SlideIn_Encaissement
    wshMenuFACT.Range("A1").Select
    
    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_Back_To_FAC_Menu()", timerStart)

End Sub

Sub FAC_Brouillon_TEC_Add_Check_Boxes(row As Long)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_TEC_Add_Check_Boxes()")
    
    Application.EnableEvents = False
    
    Dim chkBoxRange As Range: Set chkBoxRange = wshFAC_Brouillon.Range("C8:C" & row + 5)
    
    Dim cell As Range
    Dim cbx As CheckBox
        For Each cell In chkBoxRange
        ' Check if the cell is empty and doesn't have a checkbox already
        If Cells(cell.row, 8).value = False Then 'IsInvoiced = False
            'Create a checkbox linked to the cell
            Set cbx = wshFAC_Brouillon.CheckBoxes.add(cell.Left + 5, cell.Top, cell.width, cell.Height)
            With cbx
                .name = "chkBox - " & cell.row
                .value = True
                .text = ""
                .LinkedCell = cell.Address
                .Display3DShading = True
            End With
        End If
    Next cell

    With wshFAC_Brouillon
        .Range("D8:D" & row + 5).NumberFormat = "dd/mm/yyyy"
        .Range("D8:D" & row + 5).Font.Bold = False
        
        .Range("D" & row + 7).formula = "=SUMIF(C8:C" & row + 5 & ",True,G8:G" & row + 5 & ")"
        .Range("D" & row + 7).NumberFormat = "##0.00"
        .Range("D" & row + 7).Font.Bold = True
        
        .Range("B19").formula = "=SUMIF(C8:C" & row + 5 & ",True,G8:G" & row + 5 & ")"
    End With
    
    Set chkBoxRange = Nothing
    Set cell = Nothing
    Set cbx = Nothing
    
    Application.EnableEvents = True

    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_TEC_Add_Check_Boxes()", timerStart)

End Sub

Sub FAC_Brouillon_TEC_Remove_Check_Boxes(row As Long)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modFAC_Brouillon:FAC_Brouillon_TEC_Remove_Check_Boxes()")
    
    Application.EnableEvents = False
    
    Dim cbx As Shape
    For Each cbx In wshFAC_Brouillon.Shapes
        If InStr(cbx.name, "chkBox - ") Then
            cbx.delete
        End If
    Next cbx
    
    wshFAC_Brouillon.Range("C7:C" & row).value = ""  'Remove text left over
    wshFAC_Brouillon.Range("D" & row + 2).value = "" 'Remove the total formula

    Application.EnableEvents = True

    Call Output_Timer_Results("modFAC_Brouillon:FAC_Brouillon_TEC_Remove_Check_Boxes()", timerStart)

End Sub

'Sub ExportAllFacInvList() '2024-03-28 @ 14:22
'    Dim wb As Workbook
'    Dim wsSource As Worksheet
'    Dim wsTarget As Worksheet
'    Dim sourceRange As Range
'
'    Application.ScreenUpdating = False
'
'    'Work with the source range
'    Set wsSource = wshFAC_Entête
'    Dim lastUsedRow As Long
'    lastUsedRow = wsSource.Range("A99999").End(xlUp).row
'    wsSource.Range("A4:T" & lastUsedRow).Copy
'
'    'Open the target workbook
'    Workbooks.Open fileName:=wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
'                   "GCF_BD_Sortie.xlsx"
'
'    'Set references to the target workbook and target worksheet
'    Set wb = Workbooks("GCF_BD_Sortie.xlsx")
'    Set wsTarget = wb.Sheets("FACTURES")
'
'    'PasteSpecial directly to the target range
'    wsTarget.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
'    Application.CutCopyMode = False
'
'    wb.Close SaveChanges:=True
'
'    Application.ScreenUpdating = True
'
'End Sub
'
'-----------------------------------------------------------------------------------------------------------

'Sub FAC_Brouillon_Prev_PDF() '2024-03-28 @ 14:49
'
'    Call FAC_Brouillon_Goto_Onglet_FAC_Finale
'    Call FAC_Finale_Preview_PDF
'    Call FAC_Finale_Goto_Onglet_FAC_Brouillon
'
'End Sub
'
'-----------------------------------------------------------------------------------------------------------
