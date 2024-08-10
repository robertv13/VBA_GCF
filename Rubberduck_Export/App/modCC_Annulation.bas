Attribute VB_Name = "modCC_Annulation"
Option Explicit

Public invNo As String

Sub Get_Invoice_Data(noFact As String)

    'Save original worksheet
    Dim oWorkSheet As Worksheet: Set oWorkSheet = ActiveSheet
    
    'Reference to A/R master file
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).row
    
    Dim rngToSearch As Range: Set rngToSearch = ws.Range("A1").CurrentRegion.Offset(0, 0).Resize(lastUsedRow, 1)
    Dim result As Variant
    result = Application.WorksheetFunction.XLookup(noFact, _
                                                   rngToSearch, _
                                                   rngToSearch, _
                                                   "Not Found", _
                                                   0, _
                                                   1)
    
    If result <> "Not Found" Then
        'Setup the worksheet for cancellation update (kind of clipboard)
        On Error Resume Next
        Worksheets("Clipboard").Visible = xlSheetVisible
        On Error GoTo 0
        Call Erase_And_Create_Worksheet("Clipboard")
        Dim tempSheet As Worksheet: Set tempSheet = ThisWorkbook.Worksheets("Clipboard")
        Worksheets("Clipboard").Visible = xlHidden

        Dim matchedRow As Long
        matchedRow = Application.Match(noFact, rngToSearch, 0)
        Call AddRecordToTempSheet(tempSheet, ws.name, matchedRow)

        Call Display_Invoice_info(tempSheet, ws.name, matchedRow)
        
        Call Insert_Big_PDF_Icon(tempSheet)
        
        Dim resultArr As Variant
        resultArr = Fn_Get_TEC_Invoiced_By_This_Invoice(tempSheet, noFact)
        
        If Not IsEmpty(resultArr) Then
            Dim TECSummary() As Variant
            ReDim TECSummary(1 To 10, 1 To 3)
            Call Get_TEC_Summary_For_That_Invoice(tempSheet, resultArr, TECSummary)
            
            Dim FeesSummary() As Variant
            ReDim FeesSummary(1 To 5, 1 To 3)
            Call Get_Fees_Summary_For_That_Invoice(tempSheet, resultArr, FeesSummary)
            
        End If
        
        Call CC_Annulation_Get_GL_Posting(tempSheet, noFact)

        oWorkSheet.Activate
        
    Else
        MsgBox "La facture n'existe pas"
        Exit Sub
    End If
    
End Sub

Sub Insert_Big_PDF_Icon(tempSheet As Worksheet)

    Dim ws As Worksheet: Set ws = wshCC_Annulation
    
    Dim i As Long
    Dim iconPath As String
    iconPath = wshAdmin.Range("F5").value & Application.PathSeparator & "Resources\AdobeAcrobatReader.png"
    
    Dim pic As Picture
    Dim cell As Range
    
    'Loop through each row and insert the icon if there is data in column E
    Set cell = ws.Cells(7, 12) 'Set the cell where the icon should be inserted
            
    'Insert the icon
    Set pic = ws.Pictures.Insert(iconPath)
    With pic
        .Top = cell.Top + 10
        .Left = cell.Left + 10
        .Height = 50 'cell.Height
        .width = 50 'cell.width
        .Placement = xlMoveAndSize
        .OnAction = "CC_Annulation_Display_PDF_Invoice"
    End With
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    Set pic = Nothing
    Set ws = Nothing
    
End Sub

Sub CC_Annulation_Display_PDF_Invoice()

    Dim ws As Worksheet: Set ws = wshCC_Annulation
    
    'Assuming the invoice number is at 'F5'
    Dim fullPDFFileName As String
    fullPDFFileName = wshAdmin.Range("F5").value & FACT_PDF_PATH & _
        Application.PathSeparator & ws.Cells(5, 6).value & ".pdf"
    
    'Open the invoice using Adobe Acrobat Reader
    If fullPDFFileName <> "" Then
        Shell "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe " & Chr(34) & fullPDFFileName & Chr(34), vbNormalFocus
    Else
        MsgBox "Je ne retrouve pas cette facture", vbExclamation
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub

Sub Display_Invoice_info(tempSheet As Worksheet, ws As String, r As Long)

        Application.EnableEvents = False
        'Display all fields from FAC_Entête
        wshCC_Annulation.Range("L5").value = Format$(wshFAC_Entête.Cells(r, 2), "dd-mm-yyyy")
        
        wshCC_Annulation.Range("F7").value = wshFAC_Entête.Cells(r, 5)
        wshCC_Annulation.Range("F8").value = wshFAC_Entête.Cells(r, 6)
        wshCC_Annulation.Range("F9").value = wshFAC_Entête.Cells(r, 7)
        wshCC_Annulation.Range("F10").value = wshFAC_Entête.Cells(r, 8)
        wshCC_Annulation.Range("F11").value = wshFAC_Entête.Cells(r, 9)
        
        wshCC_Annulation.Range("L13").value = wshFAC_Entête.Cells(r, 10)
        wshCC_Annulation.Range("L14").value = wshFAC_Entête.Cells(r, 12)
        wshCC_Annulation.Range("L15").value = wshFAC_Entête.Cells(r, 14)
        wshCC_Annulation.Range("L16").value = wshFAC_Entête.Cells(r, 16)
        wshCC_Annulation.Range("L17").formula = "=SUM(L13:L16)"
        
        wshCC_Annulation.Range("L18").value = wshFAC_Entête.Cells(r, 18)
        wshCC_Annulation.Range("L19").value = wshFAC_Entête.Cells(r, 20)
        wshCC_Annulation.Range("L21").formula = "=SUM(L17:L19)"
        
        wshCC_Annulation.Range("L23").value = wshFAC_Entête.Cells(r, 22)
        wshCC_Annulation.Range("L25").formula = "=L21 - L23"
        
        Application.EnableEvents = True

End Sub

Sub Get_TEC_Summary_For_That_Invoice(tempSheet As Worksheet, arr As Variant, ByRef TECSummary As Variant)

    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    'Setup a Dictionary to summarize the hours by Professionnal
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")

    Dim pro As String
    Dim hres As Double
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        pro = wsTEC.Cells(arr(i), 3).value
        hres = wsTEC.Cells(arr(i), 8).value
        If hres <> 0 Then
            If dictHours.Exists(pro) Then
                dictHours(pro) = dictHours(pro) + hres
            Else
                dictHours.add pro, hres
            End If
        End If
    Next i
    
    Dim profID As Long
    Dim rowInWorksheet As Long: rowInWorksheet = 13
    Dim prof As Variant
    For Each prof In Fn_Sort_Dictionary_By_Value(dictHours, True) 'Sort dictionary by hours in descending order
        Dim strProf As String
        strProf = prof
        profID = Fn_GetID_From_Initials(strProf)
        hres = dictHours(prof)
        Dim tauxHoraire As Currency
        tauxHoraire = Fn_Get_Hourly_Rate(profID, wshCC_Annulation.Range("L5").value)
        wshCC_Annulation.Cells(rowInWorksheet, 6) = strProf
        wshCC_Annulation.Cells(rowInWorksheet, 7) = hres
        wshCC_Annulation.Cells(rowInWorksheet, 8) = tauxHoraire
        rowInWorksheet = rowInWorksheet + 1
'        Debug.Print "Summary : " & strProf & " = " & hres & " @ " & tauxHoraire
'        Cells(rowSelected, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
'        rowSelected = rowSelected + 1
    Next prof
    
    'Cleanup - 2024-07-25 @ 18:06
    Set dictHours = Nothing
    Set wsTEC = Nothing
    
End Sub

Sub Get_Fees_Summary_For_That_Invoice(tempSheet As Worksheet, arr As Variant, ByRef FeesSummary As Variant)

    Dim wsFees As Worksheet: Set wsFees = wshFAC_Sommaire_Taux
    
    'Determine the last used row
    Dim lastUsedRow As Long
    lastUsedRow = wsFees.Range("A9999").End(xlUp).row
    
    'Get Invoice number
    Dim invNo As String
    invNo = wshCC_Annulation.Range("F5").value
    
    'Use Range.Find to locate the first cell with the InvoiceNo
    Dim cell As Range
    Set cell = wsFees.Range("A2:A" & lastUsedRow).Find(What:=invNo, LookIn:=xlValues, lookAt:=xlWhole)
    
    'Check if the invNo was found at all
    Dim firstAddress As String
    Dim rowFeesSummary As Long: rowFeesSummary = 20
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Application.EnableEvents = False
        Do
            'Display values in the worksheet
            wshCC_Annulation.Range("F" & rowFeesSummary).value = wsFees.Cells(cell.row, 3).value
            wshCC_Annulation.Range("G" & rowFeesSummary).value = wsFees.Cells(cell.row, 4).value
            wshCC_Annulation.Range("H" & rowFeesSummary).value = wsFees.Cells(cell.row, 5).value
            Call AddRecordToTempSheet(tempSheet, wsFees.name, cell.row)
            rowFeesSummary = rowFeesSummary + 1
            'Find the next cell with the invNo
            Set cell = wsFees.Range("A2:A" & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        Application.EnableEvents = True
    End If
    
    'Cleanup - 2024-07-25 @ 18:06
    Set cell = Nothing
    Set wsFees = Nothing
    
End Sub

Sub CC_Annulation_Clear_Cells_And_PDF_Icon()

    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshCC_Annulation
    
    ws.Range("B3:B17, B21:B35, A38:B52").ClearContents
    
    ws.Range("F5,L5").ClearContents
    
    ws.Range("F7:I11").ClearContents
    
    ws.Range("L13:L19").ClearContents
    
    ws.Range("L21,L23,L25").ClearContents
    
    ws.Range("F13:H17").ClearContents
    
    ws.Range("F20:H24").ClearContents
    
    Dim pic As Picture
    For Each pic In ws.Pictures
        pic.delete
    Next pic
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set pic = Nothing
    Set ws = Nothing

    Application.EnableEvents = True
    
    wshCC_Annulation.Range("F5").Select
    
End Sub

Sub CC_Annulation_OK_Button_Click()

    Dim ws As Worksheet: Set ws = wshCC_Annulation
    
    Call CC_Annulation_HideButtons
    
    Call CC_Annulation_Clear_Cells_And_PDF_Icon
    
    ws.Range("F5").Select
    
    'Cleanup - 2024-07-26 @ 00:55
    Set ws = Nothing
    
End Sub

Sub CC_Annulation_Delete_Button_Click()

    Dim ws As Worksheet: Set ws = wshCC_Annulation
    
    Dim invNo As String
    invNo = ws.Range("F5").value
    
    Call CC_Annulation_HideButtons
    
    Dim answerYesNo As Long
    answerYesNo = MsgBox("Êtes-vous certain de vouloir ANNULER cette facture ? ", _
                         vbYesNo + vbQuestion, "Confirmation d'ANNULATION de facture")
    If answerYesNo = vbNo Then
        MsgBox _
            Prompt:="Cette facture ne sera PAS DÉTRUITE ! ", _
            Title:="Confirmation d'annulation", _
            Buttons:=vbCritical
            GoTo Clean_Exit
    End If
    
    If answerYesNo = vbYes Then
        Call CC_Annulation_Annule_Facture(invNo)
        
        MsgBox "La facture a été annulée" & vbNewLine & vbNewLine & _
                "Cependant le numéro est perdu à jamais", vbInformation
        
    End If
    
Clean_Exit:

    Call CC_Annulation_Clear_Cells_And_PDF_Icon

    wshCC_Annulation.Range("F5").Select
    
    'Cleanup - 2024-07-26 @ 00:55
    Set ws = Nothing
    
End Sub

Sub CC_Annulation_Get_GL_Posting(tempSheet As Worksheet, invNo)

    Dim wsGL As Worksheet: Set wsGL = wshGL_Trans
    
    Dim lastUsedRow
    lastUsedRow = wsGL.Range("A99999").End(xlUp).row
    Dim rngToSearch As Range: Set rngToSearch = wsGL.Range("D1:D" & lastUsedRow)
    
    'Use Range.Find to locate the first cell with the invNo
    Dim cell As Range
    Set cell = wsGL.Range("D2:D" & lastUsedRow).Find(What:="FACT-" & invNo, LookIn:=xlValues, lookAt:=xlWhole)
    
    'Check if the invNo was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Dim r As Long
        r = 38
        Application.EnableEvents = False
        Do
            'Save the information for invoice deletion
            Call AddRecordToTempSheet(tempSheet, wsGL.name, cell.row)
            r = r + 1
            'Find the next cell with the invNo
            Set cell = wsGL.Range("D2:D" & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        Application.EnableEvents = True
    End If

End Sub

Sub AddRecordToTempSheet(tempSheet As Worksheet, worksheetData As String, s1 As Long)

    'Find the next available row in the temporary worksheet
    Dim nextRow As Long
    With tempSheet
        If Application.WorksheetFunction.CountA(.Cells) = 0 Then
            'If the sheet is empty, start from row 1
            nextRow = 1
        Else
            'Find the last row with data and move to the next row
            nextRow = .Cells(.rows.count, 1).End(xlUp).row + 1
        End If
        
        'Add the record to the next available row
        .Cells(nextRow, 1).value = worksheetData
        .Cells(nextRow, 2).value = s1
    End With
    
End Sub

Sub CC_Annulation_Annule_Facture(invNo As String)

    MsgBox "Code à ajouter pour annuler la facture '" & invNo & "'"

End Sub
Sub CC_Annulation_ShowButtons()

    'Show the OK and CANCEL buttons
    wshCC_Annulation.Shapes("CC_Annulation_OK_Button").Visible = True
    wshCC_Annulation.Shapes("CC_Annulation_DELETE_Button").Visible = True
    
End Sub

Sub CC_Annulation_HideButtons()

    'Hide the OK and CANCEL buttons
    wshCC_Annulation.Shapes("CC_Annulation_OK_Button").Visible = False
    wshCC_Annulation.Shapes("CC_Annulation_DELETE_Button").Visible = False
    
End Sub


