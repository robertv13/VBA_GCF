Attribute VB_Name = "modCC_Annulation"
Option Explicit

Sub Get_Invoice(noFact As String)

    noFact = Trim(noFact)
    
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
        Dim matchedRow As Long
        matchedRow = Application.Match(noFact, rngToSearch, 0)
        wshCC_Annulation.Range("K5").value = Format(ws.Cells(matchedRow, 2), "dd-mm-yyyy")
        wshCC_Annulation.Range("F7").value = ws.Cells(matchedRow, 5)
        wshCC_Annulation.Range("F8").value = ws.Cells(matchedRow, 6)
        wshCC_Annulation.Range("F9").value = ws.Cells(matchedRow, 7)
        wshCC_Annulation.Range("F10").value = ws.Cells(matchedRow, 8)
        wshCC_Annulation.Range("F11").value = ws.Cells(matchedRow, 9)
        
        wshCC_Annulation.Range("K12").value = ws.Cells(matchedRow, 10)
        wshCC_Annulation.Range("K13").value = ws.Cells(matchedRow, 12)
        wshCC_Annulation.Range("K14").value = ws.Cells(matchedRow, 14)
        wshCC_Annulation.Range("K15").value = ws.Cells(matchedRow, 16)
        wshCC_Annulation.Range("K16").formula = "=SUM(K12:K15)"
        
        wshCC_Annulation.Range("K17").value = ws.Cells(matchedRow, 18)
        wshCC_Annulation.Range("K18").value = ws.Cells(matchedRow, 20)
        wshCC_Annulation.Range("K20").formula = "=SUM(K16:K18)"
        
        wshCC_Annulation.Range("K22").value = ws.Cells(matchedRow, 22)
        wshCC_Annulation.Range("K24").formula = "=SUM(K20,K22)"
        
        Call Insert_Big_PDF_Icon
        
    Else
        MsgBox "La facture n'existe pas"
        Exit Sub
    End If
    
End Sub

Sub Insert_Big_PDF_Icon()

    Dim ws As Worksheet: Set ws = wshCC_Annulation
    
    Dim i As Long
    Dim iconPath As String
    iconPath = "C:\VBA\GC_FISCALITÉ\Resources\AdobeAcrobatReader.png"
    
    Dim pic As Picture
    Dim cell As Range
    
    'Loop through each row and insert the icon if there is data in column E
    Set cell = ws.Cells(11, 7) 'Set the cell where the icon should be inserted
            
    'Insert the icon
    Set pic = ws.Pictures.Insert(iconPath)
    With pic
        .Top = cell.Top + 1
        .Left = cell.Left + 5
        .Height = cell.Height
        .width = cell.width
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
    
    Dim rowNumber As Long
    Dim fullPDFFileName As String
    
    'Determine which icon was clicked and get the corresponding row number
    Dim targetCell As Range
    Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    rowNumber = targetCell.row
    
    'Assuming the invoice number is in column E (5th column)
    fullPDFFileName = wshAdmin.Range("FolderPDFInvoice").value & _
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


